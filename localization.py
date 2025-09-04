"""
Localization Manager for Temperature Monitor
Handles loading and using language files
"""

import os
import yaml
import configparser
import locale
from pathlib import Path

class LocalizationManager:
    def __init__(self, config_file="config.ini", lang_folder="lang"):
        self.config_file = config_file
        self.lang_folder = lang_folder
        self.current_language = None
        self.translations = {}

        # Load configuration
        self.config = configparser.ConfigParser()
        self.load_config()

        # Detect and set language
        self.detect_language()

        # Load translations
        self.load_translations()

    def load_config(self):
        """Load configuration from INI file"""
        if os.path.exists(self.config_file):
            self.config.read(self.config_file, encoding='utf-8')
        else:
            # Create default config
            self.config['General'] = {'language': 'auto'}
            self.config['Arduino'] = {'port': ''}
            self.config['Origin'] = {'enabled': 'false'}
            self.config['Files'] = {'folder': ''}
            self.save_config()

    def save_config(self):
        """Save configuration to INI file"""
        with open(self.config_file, 'w', encoding='utf-8') as f:
            self.config.write(f)

    def detect_language(self):
        """Detect language from config or system locale"""
        language_setting = self.config.get('General', 'language', fallback='auto')

        if language_setting == 'auto':
            # Detect from system locale
            try:
                system_locale = locale.getdefaultlocale()[0]
                if system_locale:
                    # Map locale codes to our language codes
                    locale_lang = system_locale.lower()
                    if locale_lang.startswith('ru'):
                        self.current_language = 'ru'
                    elif locale_lang.startswith('fr'):
                        self.current_language = 'fr'
                    elif locale_lang.startswith('it'):
                        self.current_language = 'it'
                    elif locale_lang.startswith('de'):
                        self.current_language = 'de'
                    elif locale_lang.startswith('zh'):
                        self.current_language = 'zh'
                    else:
                        self.current_language = 'en'
                else:
                    self.current_language = 'en'
            except:
                self.current_language = 'en'
        else:
            self.current_language = language_setting

    def load_translations(self):
        """Load translations from YAML files"""
        lang_file = os.path.join(self.lang_folder, f"{self.current_language}.yaml")

        try:
            if os.path.exists(lang_file):
                with open(lang_file, 'r', encoding='utf-8') as f:
                    self.translations = yaml.safe_load(f)
            else:
                # Fallback to English if current language file doesn't exist
                fallback_file = os.path.join(self.lang_folder, "en.yaml")
                if os.path.exists(fallback_file):
                    with open(fallback_file, 'r', encoding='utf-8') as f:
                        self.translations = yaml.safe_load(f)
                else:
                    self.translations = {}
        except Exception as e:
            print(f"Error loading language file: {e}")
            self.translations = {}

    def get(self, key, **kwargs):
        """Get translated string with optional formatting"""
        if key in self.translations:
            text = self.translations[key]
            if kwargs:
                try:
                    return text.format(**kwargs)
                except:
                    return text
            return text
        else:
            # Return key name if translation not found
            return f"[{key}]"

    def set_language(self, language):
        """Change language (requires restart)"""
        self.current_language = language
        self.config.set('General', 'language', language)
        self.save_config()
        self.load_translations()

    def get_available_languages(self):
        """Get list of available languages"""
        languages = []
        if os.path.exists(self.lang_folder):
            for file in os.listdir(self.lang_folder):
                if file.endswith('.yaml'):
                    languages.append(file[:-5])  # Remove .yaml extension
        return languages

    def get_config(self, section, key, fallback=''):
        """Get configuration value"""
        return self.config.get(section, key, fallback=fallback)

    def set_config(self, section, key, value):
        """Set configuration value"""
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, key, str(value))
        self.save_config()

# Global instance
localization = LocalizationManager()

# Convenience function
def _(key, **kwargs):
    """Get translated string"""
    return localization.get(key, **kwargs)
