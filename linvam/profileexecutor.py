import itertools
import json
import os
import random
import re
import shlex
import signal
import subprocess
import threading
import time

import sounddevice
from vosk import Model, KaldiRecognizer

from linvam import keyboard, mouse
from linvam.keyboard import nixkeyboard as _os_keyboard
from linvam.mouse import ButtonEvent
from linvam.mouse import nixmouse as _os_mouse
from linvam.soundfiles import SoundFiles
from linvam.util import (get_language_code, get_voice_packs_folder_path, get_language_name, YDOTOOLD_SOCKET_PATH,
                         KEYS_SPLITTER, save_to_commands_file, is_push_to_listen, get_push_to_listen_hotkey, Command,
                         Config, find_best_vosk_model)

COMMAND_NAME_COMMA_SPLIT_REGGEX = r",(?![^\[]*\])"

# pylint: disable=too-many-locals
def _expand_optional_brackets(text):
    """
    Expands optional words in square brackets to create all variations.
    Commas within brackets create alternatives.

    Examples:
        "power up [the ship]" -> ["power up the ship", "power up"]
        "power up [the ship, the engines]" -> ["power up the ship", "power up the engines", "power up"]
        # pylint: disable=line-too-long
        "open [the, a] menu [now]" -> ["open the menu now", "open a menu now", "open the menu", "open a menu", "open menu now", "open menu"]

    Args:
        text: Command text potentially containing optional sections in square brackets

    Returns:
        List of all possible command variations
    """
    # Find all bracketed sections and their positions
    bracket_pattern = re.compile(r'\[([^\]]+)\]')
    matches = list(bracket_pattern.finditer(text))

    if not matches:
        # No brackets found, return original text
        return [text]

    # Extract the parts: text segments and optional segments with alternatives
    segments = []
    last_end = 0

    for match in matches:
        # Add the text before this bracket
        if match.start() > last_end:
            segments.append(('fixed', text[last_end:match.start()]))

        # Add the optional content (without brackets)
        # Split by comma to get alternatives within the bracket
        bracket_content = match.group(1)
        alternatives = [alt.strip() for alt in bracket_content.split(',')]

        # For optional brackets, we have: None (excluded) + all alternatives
        segments.append(('optional', alternatives))
        last_end = match.end()

    # Add any remaining text after the last bracket
    if last_end < len(text):
        segments.append(('fixed', text[last_end:]))

    # Build list of choices for each bracket: [None] + alternatives
    # For each optional segment, we can choose None or any of its alternatives
    optional_segments = [seg for seg in segments if seg[0] == 'optional']

    # Create choice lists: each optional bracket has [None] + [alt1, alt2, ...]
    choice_lists = []
    for seg_type, alternatives in optional_segments:
        choices = [None] + alternatives  # None means "don't include"
        choice_lists.append(choices)

    # Generate all combinations using cartesian product
    variations = []
    for combination in itertools.product(*choice_lists):
        # Build this variation
        parts = []
        optional_idx = 0

        for seg_type, content in segments:
            if seg_type == 'fixed':
                parts.append(content)
            else:  # optional
                choice = combination[optional_idx]
                if choice is not None:  # Include this choice
                    parts.append(choice)
                optional_idx += 1

        # Join and clean up extra whitespace
        variation = ''.join(parts)
        variation = ' '.join(variation.split())  # Normalize whitespace
        if variation:  # Only add non-empty variations
            variations.append(variation)

    return variations if variations else [text]


def _execute_external_command(cmd_name, is_async):
    if is_async:
        # pylint: disable=consider-using-with
        subprocess.Popen(cmd_name, shell=True)
    else:
        subprocess.run(cmd_name, shell=True, check=False)


class ProfileExecutor(threading.Thread):

    def __init__(self, p_parent=None):

        super().__init__()
        self.m_profile = None
        self.ptl_key = None
        self.ptl_keyboard_listener = None
        self.ptl_mouse_listener = None
        self.listening = False
        self.commands_list = []
        self.m_cmd_threads = {}
        self.p_parent = p_parent
        self.ydotoold = None
        if self.p_parent.m_config[Config.USE_YDOTOOL]:
            self.start_ydotoold()

        self.m_stream = None

        device_info = sounddevice.query_devices(kind='input')
        # sounddevice expects an int, sounddevice provides a float:
        self.samplerate = int(device_info['default_samplerate'])

        self.recognizer = None
        self.grammar_supported = False
        self.grammar_warning_shown = False

        self.m_sound = SoundFiles()

    def set_sound_playback_volume(self, volume):
        self.m_sound.set_volume(volume)

    def start_ydotoold(self):
        command = 'ydotoold -p ' + YDOTOOLD_SOCKET_PATH + ' -P 0666'
        args = shlex.split(command)
        # noinspection PyBroadException
        try:
            # pylint: disable=consider-using-with
            self.ydotoold = subprocess.Popen(args)
        except Exception as e:
            print('Failed to start ydotoold: ' + str(e))

    # noinspection PyUnusedLocal
    # pylint: disable=unused-argument
    def listen_callback(self, in_data, frame_count, time_info, status):
        result_string, _ = self.get_listen_result(in_data)
        if not result_string:
            return
        # Filter out ignored single words (false positives from background noise)
        if self._is_ignored_single_word(result_string):
            return
        self.check_commands(result_string)

    # noinspection PyUnusedLocal
    # pylint: disable=unused-argument
    def listen_callback_debug(self, in_data, frame_count, time_info, status):
        result_string, result_detail = self.get_listen_result(in_data)
        if not result_string:
            return

        # Filter out ignored single words (false positives from background noise)
        if self._is_ignored_single_word(result_string):
            print(f'[Ignored single word: "{result_string}"]')
            return

        # Show detailed word-level confidence if available
        if result_detail:
            print(f'Recognized: {result_string}')
            if 'result' in result_detail:
                words = result_detail['result']
                print('Word-level confidence:')
                for word_info in words:
                    word = word_info.get('word', '')
                    conf = word_info.get('conf', 0.0)
                    print(f'  {word}: {conf:.2f}')
            # Show warning if grammar should have prevented this but didn't
            if not self.grammar_supported and not self.grammar_warning_shown:
                # Check if any words in the result aren't in any commands
                recognized_words = set(result_string.split())
                all_command_words = set()
                for cmd in self.commands_list:
                    all_command_words.update(cmd.split())
                unexpected_words = recognized_words - all_command_words
                if unexpected_words:
                    print(f'⚠ Words not in command list: {unexpected_words}')
                    print('  (This is why you need grammar constraint support!)')
                    self.grammar_warning_shown = True
        else:
            print(str(result_string))
        self.check_commands(result_string)

    def _is_ignored_single_word(self, result_string):
        """Check if result is a single word in the ignore list (common false positives)."""
        # Only check single words
        words = result_string.strip().split()
        if len(words) != 1:
            return False

        # Get ignored words from config
        ignored_words = self.p_parent.m_config.get(Config.IGNORED_SINGLE_WORDS, [])
        if not ignored_words:
            # Use default if not configured
            ignored_words = ['the', 'a', 'an', 'but', 'their', 'there', 'they', 'be', 'to', 'of', 'and']

        return result_string.lower() in ignored_words

    def check_commands(self, result_string):
        for command in self.commands_list:
            if command in result_string:
                self.recognizer.Result()
                print('Detected: ' + command)
                self._do_command(command)
                break

    def get_listen_result(self, in_data):
        if self.recognizer is None:
            return '', None
        # Only process final results (when utterance is complete)
        # This ensures grammar constraints are applied and reduces spam
        # AcceptWaveform returns True when endpoint (silence) is detected
        if self.recognizer.AcceptWaveform(bytes(in_data)):
            result = self.recognizer.Result()
            result_json = json.loads(result)
            result_string = result_json.get('text', '')
            if result_string:
                if self.p_parent.m_config[Config.DEBUG]:
                    print('[Endpoint detected - finalizing utterance]')
                return result_string, result_json
        return '', None

    def set_language(self, language):
        self._stop()
        language_code = get_language_code(language)
        if language_code is None:
            print('Unsupported language: ' + language)
            return
        print('Language: ' + get_language_name(language))

        # Try to find the best available model (prefers lgraph for grammar support)
        model_path = find_best_vosk_model(language_code)

        if model_path:
            # Load specific model by path
            print(f'Loading model from: {model_path}')
            # Check if we should use faster endpoint detection
            self._configure_endpoint_detection(model_path)
            self.recognizer = KaldiRecognizer(Model(model_path=model_path), self.samplerate)
        else:
            # Fall back to auto-detection
            print(f'Auto-detecting model for language code: {language_code}')
            self.recognizer = KaldiRecognizer(Model(lang=language_code), self.samplerate)

        # Enable word-level details for better debugging
        self.recognizer.SetWords(True)

        # Apply grammar constraint if we have a profile loaded
        # This needs to happen AFTER recognizer is created
        if self.commands_list:
            self._apply_grammar_constraint()

    def _configure_endpoint_detection(self, model_path):
        """Configure faster endpoint detection for quicker response."""
        conf_path = os.path.join(model_path, 'conf', 'model.conf')
        if not os.path.exists(conf_path):
            return

        # Read current config
        with open(conf_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        # Check if already configured
        if any('LinVAM-configured' in line for line in lines):
            return  # Already configured

        # Modify endpoint rules for faster detection (shorter silence tolerance)
        # Default: rule2=0.5, rule3=1.0, rule4=2.0
        # New: rule2=0.3, rule3=0.6, rule4=0.9 (more responsive)
        modified = False
        for i, line in enumerate(lines):
            if 'endpoint.rule2.min-trailing-silence' in line:
                lines[i] = '--endpoint.rule2.min-trailing-silence=0.3\n'
                modified = True
            elif 'endpoint.rule3.min-trailing-silence' in line:
                lines[i] = '--endpoint.rule3.min-trailing-silence=0.6\n'
                modified = True
            elif 'endpoint.rule4.min-trailing-silence' in line:
                lines[i] = '--endpoint.rule4.min-trailing-silence=0.9\n'
                modified = True

        if modified:
            # Add marker comment
            lines.append('# LinVAM-configured for faster endpoint detection\n')

            # Write back
            with open(conf_path, 'w', encoding='utf-8') as f:
                f.writelines(lines)
            print('✓ Configured faster endpoint detection (0.3/0.6/0.9s silence thresholds)')

    def _init_stream(self):
        if self.recognizer is None:
            return
        if self.p_parent.m_config[Config.DEBUG]:
            callback = self.listen_callback_debug
        else:
            callback = self.listen_callback
        self.m_stream = sounddevice.RawInputStream(
            samplerate=self.samplerate,
            dtype="int16",
            channels=1,
            blocksize=4000,
            callback=callback
        )

    def _start_stream(self):
        if self.m_stream is None:
            self._init_stream()
        self.m_stream.start()

    def set_profile(self, p_profile):
        # Stop listening before changing profile to avoid crashes when applying grammar
        was_listening = self.listening
        if was_listening:
            self._stop()

        self.m_profile = p_profile
        self.commands_list = []
        if self.m_profile is None:
            print('Clearing profile')
            return
        w_commands = self.m_profile['commands']
        for w_command in w_commands:
            parts = re.split(COMMAND_NAME_COMMA_SPLIT_REGGEX, w_command['name'].strip().lower())
            for part in parts:
                # Expand optional brackets to create all variations
                variations = _expand_optional_brackets(part.strip())
                for variation in variations:
                    self.commands_list.append(variation)
        print('Profile: ' + self.m_profile['name'])
        save_to_commands_file(self.commands_list)
        # this is a dirty fix until the whole keywords recognition is refactored
        self.commands_list.sort(key=len, reverse=True)

        # Apply grammar constraint to improve accuracy
        # Must be done while stream is stopped!
        self._apply_grammar_constraint()

        # Resume listening if we were listening before
        if was_listening:
            self._start_stream()
            print('Detection resumed')

    def _apply_grammar_constraint(self):
        """Apply grammar constraint to limit recognition to command vocabulary only."""
        if self.recognizer is None:
            if self.p_parent.m_config[Config.DEBUG]:
                print('[DEBUG] Skipping grammar constraint: recognizer not initialized yet')
            return
        if not self.commands_list:
            if self.p_parent.m_config[Config.DEBUG]:
                print('[DEBUG] Skipping grammar constraint: no commands loaded yet')
            return

        # Build a grammar JSON with all command phrases
        # SetGrammar expects a JSON array of phrases
        grammar = json.dumps(self.commands_list, ensure_ascii=False)

        # Test if SetGrammar is supported
        try:
            # Check if SetGrammar method exists first
            if not hasattr(self.recognizer, 'SetGrammar'):
                self._show_grammar_not_supported_warning('SetGrammar method not available in this VOSK version')
                return

            # Try to apply grammar
            self.recognizer.SetGrammar(grammar)

            # SetGrammar might silently fail on some models
            # The only way to know if it worked is to test it
            # For now, we'll assume it worked if no exception was raised
            # The debug output will show if unexpected words are recognized
            self.grammar_supported = True

            print('\n' + '='*70)
            print(f'  GRAMMAR CONSTRAINT APPLIED: {len(self.commands_list)} command variations')
            print('='*70)
            print('⚠ NOTE: Some lgraph models accept SetGrammar but ignore it.')
            print('  Watch debug output for words NOT in your command list.')
            print('  If you see unexpected words, grammar is NOT working.')
            print('='*70 + '\n')

        except AttributeError as e:
            # SetGrammar method doesn't exist
            self._show_grammar_not_supported_warning(f'SetGrammar method not available: {e}')
        except Exception as e:
            # Some models don't support SetGrammar (static graph models)
            self._show_grammar_not_supported_warning(str(e))

    def _show_grammar_not_supported_warning(self, reason):
        """Show detailed warning when grammar constraints aren't supported."""
        self.grammar_supported = False
        print('\n' + '!'*70)
        print('  WARNING: GRAMMAR CONSTRAINT NOT WORKING')
        print('!'*70)
        print(f'Reason: {reason}')
        print()
        print('Your current VOSK model does NOT support grammar constraints.')
        print('This means it will try to recognize ANY English words, not just')
        print('your defined commands, which causes accuracy problems.')
        print()
        print('SMALL MODELS (vosk-model-small-*) DO NOT SUPPORT GRAMMAR.')
        print()
        print('To fix this, download a model that supports dynamic grammar:')
        print()
        print('  cd ~/.cache/vosk')
        print('  wget https://alphacephei.com/vosk/models/vosk-model-en-us-0.22-lgraph.zip')
        print('  unzip vosk-model-en-us-0.22-lgraph.zip')
        print()
        print('Then restart LinVAM.')
        print('!'*70 + '\n')

    def reset_listening(self):
        if self.listening:
            self.set_enable_listening(False)
            self.set_enable_listening(True)

    def set_enable_listening(self, p_enable):
        if self.recognizer is None:
            return
        if not self.listening and p_enable:
            ptl_hotkey = get_push_to_listen_hotkey()
            if is_push_to_listen() and ptl_hotkey:
                self._init_stream()
                print('Stream initialized, press ' + ptl_hotkey.name.upper() + ' to listen for commands')
                self.listening = True
                self._start_ptl(ptl_hotkey)
            else:
                self._start_stream()
                print('Detection started')
                if not self.grammar_supported and self.p_parent.m_config[Config.DEBUG]:
                    print('[DEBUG] Note: grammar_supported flag is False')
                    print('[DEBUG] This may be incorrect - check for grammar constraint messages above')
                self.listening = True
        elif self.listening and not p_enable:
            self._stop()

    def _start_ptl(self, ptl_hotkey):
        self.ptl_key = ptl_hotkey
        if ptl_hotkey.is_mouse_key:
            self.ptl_mouse_listener = mouse.hook(self._on_mouse_key_event)
        else:
            self.ptl_keyboard_listener = keyboard.hook(self._on_keyboard_key_event)

    def _on_mouse_key_event(self, event):
        if not isinstance(event, ButtonEvent):
            return
        if str(event.button) == str(self.ptl_key.button):
            if event.event_type == mouse.DOWN and not self.m_stream.active:
                self.m_stream.start()
            elif event.event_type == mouse.UP and self.m_stream.active:
                self._stop_ptl_stream()

    def _stop_ptl_stream(self):
        # sleep for 1 second to allow said commands to be processed correctly
        time.sleep(0.5)
        self.m_stream.stop()
        self.recognizer.Result()

    def _on_keyboard_key_event(self, event):
        if event.name == 'unknown':
            return
        if int(event.scan_code) == int(self.ptl_key.code):
            if event.event_type == keyboard.KEY_DOWN and not self.m_stream.active:
                self.m_stream.start()
            elif event.event_type == keyboard.KEY_UP and self.m_stream.active:
                self._stop_ptl_stream()

    def _stop(self):
        if self.m_stream is not None:
            self._stop_ptl_listener()
            self.m_stream.stop()
            self.m_stream.close()
            self.m_stream = None
            self.recognizer.FinalResult()
            self.listening = False
            print('Detection stopped')

    def shutdown(self):
        # Stop all running command threads
        for cmd_name in list(self.m_cmd_threads.keys()):
            self._stop_command(cmd_name)

        self.m_sound.stop()
        self._stop()
        if self.ydotoold is not None:
            try:
                os.kill(self.ydotoold.pid, signal.SIGKILL)
                self.ydotoold = None
            except OSError:
                pass

    def _stop_ptl_listener(self):
        # noinspection PyBroadException
        # pylint: disable=bare-except,R0801
        try:
            if self.ptl_keyboard_listener is not None:
                self.ptl_keyboard_listener()
                self.ptl_keyboard_listener = None
            if self.ptl_mouse_listener is not None:
                mouse.unhook(self.ptl_mouse_listener)
                self.ptl_mouse_listener = None
        except Exception as ex:
            print(str(ex))

    def do_action(self, p_action):
        # {'name': 'key action', 'key': 'left', 'type': 0}
        # {'name': 'pause action', 'time': 0.03}
        # {'name': 'command stop action', 'command name': 'down'}
        # {'name': 'mouse move action', 'x':5, 'y':0, 'absolute': False}
        # {'name': 'mouse click action', 'button': 'left', 'type': 0}
        # {'name': 'mouse wheel action', 'delta':10}
        w_action_name = p_action['name']
        match w_action_name:
            case Command.KEY_ACTION:
                self._press_key(p_action)
            case Command.PAUSE_ACTION:
                print("Sleep ", p_action['time'])
                time.sleep(p_action['time'])
            case Command.COMMAND_STOP_ACTION:
                self._stop_command(p_action['command name'])
            case Command.EXECUTE_VOICE_COMMAND_ACTION:
                self._execute_voice_command(p_action['command name'])
            case Command.EXECUTE_EXTERNAL_COMMAND_ACTION:
                _execute_external_command(p_action['command'], False)
            case Command.COMMAND_PLAY_SOUND | Command.PLAY_SOUND:
                self._play_sound(p_action)
            case Command.STOP_SOUND:
                self.m_sound.stop()
            case Command.MOUSE_MOVE_ACTION:
                self._move_mouse(p_action)
            case Command.MOUSE_CLICK_ACTION:
                self._click_mouse_key(p_action)
            case Command.MOUSE_SCROLL_ACTION:
                self._scroll_mouse(p_action)

    def _move_mouse(self, action):
        if self.p_parent.m_config[Config.USE_YDOTOOL]:
            self._move_mouse_ydotool(action)
        else:
            self._move_mouse_mouse(action)

    @staticmethod
    def _move_mouse_mouse(p_action):
        if p_action['absolute']:
            _os_mouse.move_to(p_action['x'], p_action['y'])
        else:
            _os_mouse.move_relative(p_action['x'], p_action['y'])

    def _move_mouse_ydotool(self, p_action):
        if p_action['absolute']:
            command = 'mousemove --absolute -x ' + str(p_action['x']) + " -y " + str(p_action['y'])
        else:
            command = 'mousemove -x ' + str(p_action['x']) + " -y " + str(p_action['y'])
        self._execute_ydotool_command(command)

    def _scroll_mouse(self, action):
        if self.p_parent.m_config[Config.USE_YDOTOOL]:
            self._scroll_mouse_ydotool(action)
        else:
            self._scroll_mouse_mouse(action)

    @staticmethod
    def _scroll_mouse_mouse(p_action):
        _os_mouse.wheel(int(p_action['delta']))

    def _scroll_mouse_ydotool(self, p_action):
        command = 'mousemove --wheel -x 0 -y ' + str(p_action['delta'])
        self._execute_ydotool_command(command)

    def _click_mouse_key(self, action):
        if self.p_parent.m_config[Config.USE_YDOTOOL]:
            self._click_mouse_key_ydotool(action)
        else:
            self._click_mouse_key_mouse(action)

    @staticmethod
    def _click_mouse_key_mouse(p_action):
        w_type = p_action['type']
        w_button = str(p_action['button'])
        match w_type:
            case 1:
                _os_mouse.press(w_button)
            case 0:
                _os_mouse.release(w_button)
            case 10:
                _os_mouse.press(w_button)
                _os_mouse.release(w_button)
            case 11:
                _os_mouse.press(w_button)
                _os_mouse.release(w_button)
                _os_mouse.press(w_button)
                _os_mouse.release(w_button)
            case _:
                print("Unknown mouse type " + w_type + " , skipping")

    def _click_mouse_key_ydotool(self, p_action):
        w_type = p_action['type']
        w_button = p_action['button']
        click_command = '0x'
        match w_type:
            case 1:
                click_command += '4'
            case 0:
                click_command += '8'
            case 10:
                click_command += 'C'
            case 11:
                click_command += 'C'
            case _:
                click_command += '0'

        match w_button:
            case 'left':
                click_command += '0'
            case 'middle':
                click_command += '2'
            case 'right':
                click_command += '1'
            case _:
                click_command += '0'

        args = ""
        if w_type == 11:
            args = "--repeat 2 "

        command = 'click --next-delay 100 ' + args + click_command
        self._execute_ydotool_command(command)

    def _execute_ydotool_command(self, command):
        if self.ydotoold is not None:
            os.system('env YDOTOOL_SOCKET=' + YDOTOOLD_SOCKET_PATH + ' ydotool ' + command)
            if self.p_parent.m_config[Config.DEBUG]:
                print('Executed ydotool command: ' + command)
        else:
            print('ydotoold daemon not running')

    class CommandThread(threading.Thread):
        def __init__(self, p_profile_executor, p_actions, p_repeat):
            threading.Thread.__init__(self)
            self.daemon = True  # Allow app to exit even if thread is running
            self.profile_executor = p_profile_executor
            self.m_actions = p_actions
            self.m_repeat = p_repeat
            self.m_stop = False

        def run(self):
            w_repeat = self.m_repeat
            while not self.m_stop and w_repeat > 0:
                for w_action in self.m_actions:
                    self.profile_executor.do_action(w_action)
                w_repeat -= 1

        def stop(self):
            self.m_stop = True
            # Don't join() - just signal the thread to stop
            # Since it's a daemon thread, it will be terminated when the process exits

    def _do_command(self, p_cmd_name):
        w_command = self._get_command_for_executing(p_cmd_name)
        if w_command is None:
            return
        w_actions = w_command['actions']
        w_async = w_command['async']
        if not w_async:
            w_repeat = w_command['repeat']
            w_repeat = max(w_repeat, 1)
            while w_repeat > 0:
                for w_action in w_command['actions']:
                    self.do_action(w_action)
                w_repeat -= 1
        else:
            w_cmd_thread = ProfileExecutor.CommandThread(self, w_actions, w_command['repeat'])
            w_cmd_thread.start()
            self.m_cmd_threads[p_cmd_name] = w_cmd_thread

    def _get_command_for_executing(self, cmd_name):
        if self.m_profile is None:
            return None

        w_commands = self.m_profile['commands']
        command = None
        for w_command in w_commands:
            parts = re.split(COMMAND_NAME_COMMA_SPLIT_REGGEX, w_command['name'].strip().lower())
            for part in parts:
                # Expand optional brackets to match all variations
                variations = _expand_optional_brackets(part.strip())
                for variation in variations:
                    if variation.lower() == cmd_name:
                        command = w_command
                        break
                if command is not None:
                    break

            if command is not None:
                break
        return command

    def _execute_voice_command(self, cmd_name):
        self._do_command(cmd_name)

    def _stop_command(self, p_cmd_name):
        if p_cmd_name in self.m_cmd_threads:
            self.m_cmd_threads[p_cmd_name].stop()
            del self.m_cmd_threads[p_cmd_name]

    def _play_sound(self, p_cmd_name):
        # Support both single file (backward compatibility) and multiple files with random selection
        if 'files' in p_cmd_name and p_cmd_name['files']:
            # Multiple files - randomly choose one
            selected_file = random.choice(p_cmd_name['files'])
        elif 'file' in p_cmd_name:
            # Single file (backward compatibility)
            selected_file = p_cmd_name['file']
        else:
            print("ERROR - No sound file specified in action")
            return

        sound_file = (get_voice_packs_folder_path() + p_cmd_name['pack'] + '/' + p_cmd_name['cat'] + '/'
                      + selected_file)
        self.m_sound.play(sound_file)

    def _press_key(self, action):
        if self.p_parent.m_config[Config.USE_YDOTOOL]:
            self._press_key_ydotool(action)
        else:
            self._press_key_keyboard(action)

    def _press_key_ydotool(self, action):
        events = str(action['key_events']).replace(KEYS_SPLITTER, ' ')
        self._execute_ydotool_command('key -d ' + str(action['delay']) + ' ' + events)

    @staticmethod
    def _press_key_keyboard(action):
        events = str(action['key_events']).split(KEYS_SPLITTER)
        for event in events:
            splits = event.split(':')
            code = int(splits[0])
            match splits[1]:
                case '1':
                    _os_keyboard.press(code)
                    time.sleep(int(action['delay']) / 1000)
                case '0':
                    _os_keyboard.release(code)
