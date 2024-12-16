# Import packages
import tkinter as tk
from tkinter import messagebox
from tkinter import *
import pandas as pd
import math


###################### Load powder price ######################
# Define the file path to the Excel sheet
excel_file_path = "ColourPrice.xlsx"

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
    
# Load colour prices from the Excel sheet
colour_prices = load_colours_from_excel(excel_file_path)


###################### Create searchable list for colour selection ######################
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

# Function to calculate powder cost per m^2
def powder_price(powderprice_kg):
    powderprice_m2 = powderprice_kg * 0.2
    return powderprice_m2



def space(quantity):
    conveyor_height = 1700-200
    # quantity = int(quant_entry.get())
    bound_dim = [float(bound1.get()), float(bound2.get()), float(bound3.get())]
    bound_dim.sort(reverse=True)
    hori_gap = bound_dim[2]
    vert_gap = bound_dim[2]


    if bound_dim[0] > 100 and bound_dim[1] <= conveyor_height and bound_dim[0] <= conveyor_height:
        # Calculate conveyor space when largest dimension is horisontal
        if math.floor(conveyor_height/(bound_dim[1]+vert_gap)) >= 1:
            vert_stack1 = math.floor(conveyor_height/(bound_dim[1]+vert_gap))
        else:
            vert_stack1 = 1
        hori_stack1 = math.ceil(quantity/vert_stack1)
        space1 = (bound_dim[0]+hori_gap)*hori_stack1

        # Calculate conveyor space when largest dimension is vertical
        if math.floor(conveyor_height/(bound_dim[0]+vert_gap)) >= 1:
            vert_stack2 = math.floor(conveyor_height/(bound_dim[0]+vert_gap))
        else:
            vert_stack2 = 1
        hori_stack2 = math.ceil(quantity/vert_stack2)
        space2 = (bound_dim[1]+hori_gap)*hori_stack2

        # Compare the two ways of hanging the part
        if space1 < space2:
            print("Longest dimension should be hung horisontal with ", vert_stack1, "stacked vertically")
            print("Conveyor space needed: ", space1)
            space_final = space1

        else:
            print("Longest dimension should be hung vertical with ", vert_stack2, " stacked vertically")
            print("Conveyor space needed: ", space2)
            space_final = space2
        
    elif bound_dim[0] > 100 and bound_dim[1] <= conveyor_height:
        # Calculate conveyor space when largest dimension is horisontal
        if math.floor(conveyor_height/(bound_dim[1]+vert_gap)) >= 1:
            vert_stack1 = math.floor(conveyor_height/(bound_dim[1]+vert_gap))
        else:
            vert_stack1 = 1
        hori_stack1 = math.ceil(quantity/vert_stack1)
        space1 = (bound_dim[0]+hori_gap)*hori_stack1

        print("Longest dimension should be hung horisontal")
        print("Conveyor space needed: ", space1)
        space_final = space1

    elif bound_dim[0] > conveyor_height and bound_dim[1] > conveyor_height:
        messagebox.showerror("Error", "Part is to big to fit on conveyor")
        space_final = 69

    else:
        messagebox.showerror("Error", "Parts should be arranged on a rack")
        space_final = 69

    final_space = [space_final/quantity, space_final]
    return final_space



###################### Cost Calculation ######################
def calc_price():
    try:
        # Get quantity
        quantity = int(quant_entry.get())

        # Get the selected colour and corresponding price per kg
        selected_colour = colour_var.get()
        powderprice_kg = colour_prices[selected_colour]

        # Get powder price per square meter
        ppm2 = powder_price(powderprice_kg)

        # Get surface area from input
        surface_area = float(surface.get())

        # Get Space on conveyor
        conveyor_space = space(quantity)

        # Perform the price calculation
        part_cost = round((surface_area / 1000000 * ppm2) + conveyor_space[0]/1000 * 106, 2)
        total_cost = round((surface_area / 1000000 * ppm2)*quantity + conveyor_space[1]/1000 * 106, 2)
        # Display the result in the result label
        result_label.config(text=f"Cost per part: {part_cost} DKK")
        result_label1.config(text=f"Total cost: {total_cost} DKK")

    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid numbers for surface area, bounding box, and quantity!")



###################### UI code ######################
leftmargin = 20

# Create the main window
root = tk.Tk()
root.title("Cost Calculator")
root.geometry("450x400")

# Create a label for surface area
Label(root, text="Enter Surface Area (mm²):").place(x=leftmargin, y=10)

# Create entry for surface area
surface = tk.Entry(root, width=30)
surface.place(x=leftmargin, y=30)



# Create label for bounding box
Label(root, text="Enter dimensions in mm of bounding box around the part:").place(x=leftmargin, y=60)
Label(root, text="1. dimension").place(x=leftmargin, y=80)
Label(root, text="2. dimension").place(x=leftmargin+100, y=80)
Label(root, text="3. dimension").place(x=leftmargin+200, y=80)

# Create entry for 1. dimension of bounding box
bound1 = tk.Entry(root, width=10)
bound1.place(x=leftmargin, y=100)

# Create entry for 2. dimension of bounding box
bound2 = tk.Entry(root, width=10)
bound2.place(x=leftmargin+100, y=100)

# Create entry for 3. dimension of bounding box
bound3 = tk.Entry(root, width=10)
bound3.place(x=leftmargin+200, y=100)

# Create label for quantity
Label(root, text="Enter quantity:").place(x=250, y=150)
# Create entry for quantity
quant_entry = tk.Entry(root, width=10)
quant_entry.place(x=250, y=170)

# Create a label asking for colour selection
colour_label = tk.Label(root, text="Select a powder:").place(x=leftmargin, y=150)

# Create an entry for searching colours
search_var = tk.StringVar()
search_entry = tk.Entry(root, textvariable=search_var, width=30)
search_entry.place(x=leftmargin, y=170)

# Create a listbox to show the colours
listbox = tk.Listbox(root, width=30, height=5)
listbox.place(x=leftmargin, y=200)

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
calculate_button = tk.Button(root, text="Calculate", command=calc_price).place(x=leftmargin, y=300)

# Create a label to display the result
result_label = tk.Label(root, text="Result: ")
result_label.place(x=leftmargin, y=330)
result_label1 = tk.Label(root, text="Result: ")
result_label1.place(x=leftmargin, y=350)


# Start the Tkinter event loop
root.mainloop()
