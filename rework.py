import tkinter as tk

# Create the main window
window = tk.Tk()
window.state('zoomed')  # Maximize the window
window.title("N. Gombolyítás")

# Configure the main window's grid
window.columnconfigure(0, weight=1)
window.rowconfigure(0, weight=1)
window.rowconfigure(1, weight=1)
window.rowconfigure(2, weight=1)
window.rowconfigure(3, weight=1)
window.rowconfigure(4, weight=2)

# Create Frame widgets
frame1 = tk.Frame(window)
frame2 = tk.Frame(window)

# Configure the grid of frame1
frame1.columnconfigure(0, weight=1)
for i in range(4):
    frame1.rowconfigure(i, weight=1)

# Place Frame widgets in the main window
frame1.grid(row=0, column=0, rowspan=4, sticky='nsew')
frame2.grid(row=4, column=0, sticky='nsew')

# Create a Canvas inside frame1
canvas = tk.Canvas(frame1)
canvas.grid(row=0, column=0, rowspan=4, sticky='nsew')

# Load and display an image on the Canvas
image = tk.PhotoImage(file="Ngomb_bground.png")
canvas.create_image(0, 0, anchor=tk.NW, image=image)

# Create a Label widget inside frame1
label = tk.Label(frame1, text="Ez egy label")
label.grid(row=0, column=0, sticky='nsew')  # Use grid instead of pack

# Start the Tkinter event loop
window.mainloop()
