#!/usr/bin/env python3
"""
Language Configuration Utility
Allows users to change the application language
"""

import configparser
import os
import sys

def show_available_languages():
    """Show available languages"""
    languages = {
        'en': 'English',
        'ru': 'Русский (Russian)',
        'fr': 'Français (French)',
        'it': 'Italiano (Italian)',
        'de': 'Deutsch (German)',
        'zh': '中文 (Chinese Simplified)'
    }

    print("Available languages:")
    for code, name in languages.items():
        print(f"  {code} - {name}")

def get_current_language():
    """Get current language setting"""
    config = configparser.ConfigParser()
    if os.path.exists('config.ini'):
        config.read('config.ini', encoding='utf-8')
        return config.get('General', 'language', fallback='auto')
    return 'auto'

def set_language(language):
    """Set language in config.ini"""
    config = configparser.ConfigParser()

    # Read existing config or create new
    if os.path.exists('config.ini'):
        config.read('config.ini', encoding='utf-8')

    # Ensure General section exists
    if not config.has_section('General'):
        config.add_section('General')

    # Set language
    config.set('General', 'language', language)

    # Write config
    with open('config.ini', 'w', encoding='utf-8') as f:
        config.write(f)

    print(f"Language set to: {language}")
    print("Restart the application to apply changes.")

def main():
    """Main function"""
    print("=== Temperature Monitor Language Configuration ===")
    print()

    current = get_current_language()
    print(f"Current language setting: {current}")
    print()

    show_available_languages()
    print()

    if len(sys.argv) > 1:
        # Language provided as command line argument
        language = sys.argv[1].lower()
        if language in ['en', 'ru', 'fr', 'it', 'de', 'zh', 'auto']:
            set_language(language)
        else:
            print(f"Error: Invalid language code '{language}'")
            print("Valid codes: en, ru, fr, it, de, zh, auto")
            sys.exit(1)
    else:
        # Interactive mode
        while True:
            try:
                choice = input("Enter language code (or 'q' to quit): ").lower().strip()
                if choice == 'q':
                    break
                elif choice in ['en', 'ru', 'fr', 'it', 'de', 'zh', 'auto']:
                    set_language(choice)
                    break
                else:
                    print("Invalid choice. Please enter a valid language code.")
            except KeyboardInterrupt:
                print("\nExiting...")
                break

if __name__ == "__main__":
    main()
