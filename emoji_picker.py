import json
import os
import webview
from jinja2 import Template

class EmojiPicker:
    def __init__(self, parent=None, on_emoji_selected=None):
        self.parent = parent
        self.on_emoji_selected = on_emoji_selected
        self.window = None
        
    def select_emoji(self, emoji):
        """Called from JavaScript when an emoji is selected"""
        if self.on_emoji_selected:
            self.on_emoji_selected(emoji['char'])
        if self.window:
            self.window.destroy()
    
    def show(self):
        """Show the emoji picker window"""
        # Load emoji data
        emoji_path = os.path.join(os.path.dirname(__file__), 'data', 'emoji_data.json')
        with open(emoji_path, 'r', encoding='utf-8') as f:
            emoji_data = json.load(f)
            
        # Load template
        template_path = os.path.join(os.path.dirname(__file__), 'templates', 'emoji_picker.html')
        with open(template_path, 'r', encoding='utf-8') as f:
            template = Template(f.read())
            
        # Render template with emoji data
        html = template.render(emoji_data=emoji_data)
        
        # Create window
        self.window = webview.create_window(
            'Emoji Picker',
            html=html,
            width=400,
            height=500,
            resizable=True,
            js_api=self
        )
        
        # Start window in a separate thread
        webview.start() 