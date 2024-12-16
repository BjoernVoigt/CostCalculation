# Phase 1 Price Calculation with Searchable colour Selection from Excel
import tkinter as tk
from tkinter import messagebox
import pandas as pd

# Load colours and prices from an Excel file
def load_colours_from_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        # Create a dictionary with colour as key and price as value
        colour_prices = dict(zip(df['Colour'], df['Price']))
        return colour_prices
    except Exception as e:
        print(f"Failed to load Excel file: {e}")  # Print error for debugging
        messagebox.showerror("File Error", f"Failed to load colours from file: {e}")
        return {}

# Function to update the Listbox with search results
def update_listbox(event):
    search_term = search_var.get().lower()  # Get the search term and make it lowercase
    listbox.delete(0, tk.END)  # Clear the Listbox
    
    # Loop through the available colours and add the ones that match the search term
    for colour in colour_prices.keys():
        if search_term in colour.lower():
            listbox.insert(tk.END, colour)

# Function to select a colour from the Listbox
def select_colour(event):
    selected_colour = listbox.get(listbox.curselection())
    colour_var.set(selected_colour)  # Update the colour_var with the selected colour

def powder_price(powderprice_kg):
    powderprice_m2 = powderprice_kg * 0.2
    return powderprice_m2

def cost_calc():
    try:
        # Get the selected colour and corresponding price per kg
        selected_colour = colour_var.get()
        powderprice_kg = colour_prices[selected_colour]

        # Get powder price per square meter
        ppm2 = powder_price(powderprice_kg)

        # Get surface area from input
        surface_area = float(entry.get())

        # Perform the price calculation
        price = round(surface_area / 1000000 * ppm2, 2)

        # Display the result in the result label
        result_label.config(text=f"Cost per part: {price} DKK")

    except ValueError:
        messagebox.showerror("Input Error", "Please enter a valid number for surface area!")
    except KeyError:
        messagebox.showerror("Input Error", "Please select a valid colour!")

# Define the file path to the Excel sheet
excel_file_path = "ColourPrice.xlsx"

# Load colour prices from the Excel sheet
colour_prices = load_colours_from_excel(excel_file_path)

# Create the main window
root = tk.Tk()
root.title("Cost Calculator")
root.geometry("400x350")

# Create a label for surface area
label = tk.Label(root, text="Enter Surface Area (mm²):")
label.pack(pady=10)

# Create entry for surface area
entry = tk.Entry(root, width=30)
entry.pack(pady=5)

# Create a label asking for colour selection
colour_label = tk.Label(root, text="Select a powder:")
colour_label.pack(pady=10)

# Create an entry for searching colours
search_var = tk.StringVar()
search_entry = tk.Entry(root, textvariable=search_var, width=30)
search_entry.pack(pady=5)

# Create a listbox to show the colours
listbox = tk.Listbox(root, width=30, height=5)
listbox.pack(pady=5)

# Bind the search entry and the listbox to respective functions
search_entry.bind('<KeyRelease>', update_listbox)
listbox.bind('<<ListboxSelect>>', select_colour)

# Create a variable to store the selected colour
colour_var = tk.StringVar(root)
colour_var.set("Select a colour")  # Set default value

# Populate the listbox with all available colours initially
for colour in colour_prices.keys():
    listbox.insert(tk.END, colour)

# Create a calculate button
calculate_button = tk.Button(root, text="Calculate", command=cost_calc)
calculate_button.pack(pady=10)

# Create a label to display the result
result_label = tk.Label(root, text="Result: ")
result_label.pack(pady=10)

# Start the Tkinter event loop
root.mainloop()
