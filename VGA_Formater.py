import tkinter as tk
from tkinter import scrolledtext

def format_vba_code(vba_code):
    lines = vba_code.split('\n')
    formatted_lines = []
    indent_level = 0
    indent_string = '    '  # Use 4 spaces for indentation

    for line in lines:
        stripped_line = line.strip()
        lower_stripped_line = stripped_line.lower()
        
        # Handle decrease indent for ending statements first
        if any(lower_stripped_line.startswith(end) for end in ['end sub', 'end if', 'end with', 'end select', 'next', 'loop', 'end function']):
            indent_level -= 1

        # Handle special case for else statement
        if lower_stripped_line.startswith('else'):
            formatted_line = (indent_string * (indent_level - 1)) + stripped_line
        else:
            formatted_line = (indent_string * indent_level) + stripped_line
        formatted_lines.append(formatted_line)

        # Adjust indentation level for start statements after adding the line
        if any(lower_stripped_line.startswith(start) for start in ['sub', 'function', 'if', 'for', 'while', 'do', 'select case', 'with']):
            indent_level += 1

        # Re-increase indent for else statement
        if lower_stripped_line.startswith('else'):
            indent_level += 1

    formatted_code = '\n'.join(formatted_lines)
    return formatted_code




def format_and_display():
    input_code = input_text.get("1.0", tk.END)
    formatted_code = format_vba_code(input_code)
    output_text.delete("1.0", tk.END)
    output_text.insert(tk.END, formatted_code)

# Create the main window
window = tk.Tk()
window.title("VBA Formatter")

# Create a text box for input
input_label = tk.Label(window, text="Input VBA Code:")
input_label.pack()

input_text = scrolledtext.ScrolledText(window, width=80, height=10)
input_text.pack()

# Create a button to format the code
format_button = tk.Button(window, text="Format Code", command=format_and_display)
format_button.pack()

# Create a text box for output
output_label = tk.Label(window, text="Formatted VBA Code:")
output_label.pack()

output_text = scrolledtext.ScrolledText(window, width=80, height=10)
output_text.pack()

# Start the Tkinter event loop
window.mainloop()
