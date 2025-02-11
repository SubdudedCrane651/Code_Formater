import tkinter as tk
from tkinter import scrolledtext
from bs4 import BeautifulSoup

def format_html_code(html_code):
    soup = BeautifulSoup(html_code, 'html.parser')
    formatted_html = soup.prettify()
    return formatted_html

def format_and_display():
    input_code = input_text.get("1.0", tk.END)
    formatted_code = format_html_code(input_code)
    output_text.delete("1.0", tk.END)
    output_text.insert(tk.END, formatted_code)

# Create the main window
window = tk.Tk()
window.title("HTML Formatter")

# Create a text box for input
input_label = tk.Label(window, text="Input HTML Code:")
input_label.pack()

input_text = scrolledtext.ScrolledText(window, width=80, height=10)
input_text.pack()

# Create a button to format the code
format_button = tk.Button(window, text="Format Code", command=format_and_display)
format_button.pack()

# Create a text box for output
output_label = tk.Label(window, text="Formatted HTML Code:")
output_label.pack()

output_text = scrolledtext.ScrolledText(window, width=80, height=10)
output_text.pack()

# Start the Tkinter event loop
window.mainloop()
