import tkinter as tk
from tkinter import filedialog, messagebox, ttk, font, Text, PhotoImage
from PIL import Image, ImageTk
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Importing data processing functions here
from Filter_X import DataProcessor_X
from Filter_XII import DataProcessor_XII

# Creating a tkinter window
root = tk.Tk()
root.title("Class Data Processor")
root.iconbitmap('Items/logo.ico')
window_width = root.winfo_screenwidth() - 15
window_height = root.winfo_screenheight() - 40
root.geometry("{0}x{1}+0+0".format(window_width, window_height))  # Full screen

# All Images:
width1, height1 = 128, 128
width2, height2 = 64, 64
background_image = Image.open('Items/background1.png')
back_image = PhotoImage(file='Items/back.png')
next_image = PhotoImage(file='Items/next.png')
main_image = PhotoImage(file='Items/main.png')
export_image = PhotoImage(file='Items/export.png')

about = 'Items/about.txt'

background_image = background_image.resize((window_width+15, window_height+15))
back_image = back_image.subsample(back_image.width() // width1, back_image.height() // height1)
next_image = next_image.subsample(next_image.width() // width1, next_image.height() // height1)
main_image = main_image.subsample(main_image.width() // width2, main_image.height() // height2)
export_image = export_image.subsample(export_image.width() // width2, export_image.height() // height2)

# Background
my_img=ImageTk.PhotoImage(background_image)
my_label=tk.Label(root,image=my_img)
my_label.place(x=-2, y=-2)

def select_class(class_type):       # Function to select class X or XII
    if class_type == "X":
        import_file_x()  # Call the import_file function for class X
    elif class_type == "XII":
        import_file_xii()  # Call the import_file function for class XII

def destroy_all_widgets(root):
    for widget in root.winfo_children():
        widget.destroy()

def result_tab(tab_,dafra):
    ##########################   TAB 2   #########################
    # Creating a frame for tree1 
    frame_tree1 = ttk.Frame(tab_,)
    frame_tree1.place(relheight=0.9, relwidth=0.3, relx=0, rely=0)

    # Creating a Treeview for the first 3 columns
    tree1 = ttk.Treeview(frame_tree1, columns=list(dafra.columns[:3]), show="headings")

    # Creating a frame for tree2
    frame_tree2 = ttk.Frame(tab_,)
    frame_tree2.place(relheight=0.9, relwidth=0.7, relx=0.3, rely=0)

    # Creating a Treeview for the remaining columns
    tree2 = ttk.Treeview(frame_tree2, columns=list(dafra.columns[3:]), show="headings")

    # Add a vertical scrollbar for both Treeviews
    vertical_scrollbar = ttk.Scrollbar(tab_, orient="vertical", command=lambda *args: on_scroll(*args))

    tree1.configure(yscrollcommand=vertical_scrollbar.set)
    tree2.configure(yscrollcommand=vertical_scrollbar.set)

    # Configuring column headings
    for col in dafra.columns[:3]:
        tree1.heading(col, text=col,)
        tree1.column(col, width=150)

    # Setting the width of the first 3 columns 
    column_widths = {
        dafra.columns[0]: 90,
        dafra.columns[1]: 80,
        dafra.columns[2]: 200,
        }
    for column, width in column_widths.items():
        tree1.column(column, width=width, anchor="w",)

    # Remaining columns
    for col in dafra.columns[3:]:
        tree2.heading(col, text=col)
        tree2.column(col, width=100, anchor='center', stretch=False)

    # Inserting data into the Treeviews
    for index, row in dafra.iterrows():
        values1 = [row[col] for col in dafra.columns[:3]]
        values2 = [row[col] for col in dafra.columns[3:]]
        tree1.insert("", "end", values=values1)
        tree2.insert("", "end", values=values2)

    # Coordinates layout for Treeviews and Scrollbar
    tree1.place(relheight=1, relwidth=1, rely=0)
    tree2.place(relheight=1, relwidth=1, rely=0)
    vertical_scrollbar.place(relx=1, rely=0.04, relheight=0.86, anchor="ne")

    # Function to sync the vertical scrollbar
    def on_scroll(*args):
        tree1.yview(*args)
        tree2.yview(*args)

    # Create a horizontal scrollbar for tree2 and placing it
    horizontal_scrollbar = ttk.Scrollbar(tab_, orient="horizontal", command=tree2.xview)
    tree2.configure(xscrollcommand=horizontal_scrollbar.set)
    horizontal_scrollbar.place(relx=0.3, rely=0.9, relwidth=0.7, anchor="sw")

    # # Bind the mouse scroll event to both Treeviews
    # def on_mouse_scroll(event):
    #     if tree1.yview_scroll:
    #         tree1.yview_scroll(-1 if event.delta > 0 else 1, 'units')
    #         tree2.yview_scroll(-2 if event.delta > 0 else 2, 'units')
    #     elif tree2.yview_scroll:
    #         tree2.yview_scroll(-1 if event.delta > 0 else 1, 'units')
    #         tree1.yview_scroll(-2 if event.delta > 0 else 2, 'units')
    # tree1.bind("<MouseWheel>", on_mouse_scroll)
    # tree2.bind("<MouseWheel>", on_mouse_scroll)

    def on_arrow_scroll(event):
        if event.keysym in ('Up', 'Down'):
            tree1.yview_scroll(-1 if event.keysym == 'Up' else 1, 'units')
            tree2.yview_scroll(-1 if event.keysym == 'Up' else 1, 'units')

    tree1.bind("<Up>", on_arrow_scroll)
    tree1.bind("<Down>", on_arrow_scroll)
    tree2.bind("<Up>", on_arrow_scroll)
    tree2.bind("<Down>", on_arrow_scroll)

    # Disable mouse scroll for both Treeviews
    def disable_mouse_scroll(event):
        return "break"

    tree1.bind("<MouseWheel>", disable_mouse_scroll)
    tree2.bind("<MouseWheel>", disable_mouse_scroll)

    # Configure grid weights to make the Treeviews and scrollbar expand with the window
    frame_tree1.grid_rowconfigure(0, weight=1)
    frame_tree1.grid_columnconfigure(0, weight=1)
    frame_tree2.grid_rowconfigure(0, weight=1)
    frame_tree2.grid_columnconfigure(0, weight=1)

def analysis1_tab(tab_, dafra):
    # Creating a frame to contain the Treeview for tab2
    frame_tab2 = tk.Frame(tab_, height=900)
    frame_tab2.pack(fill="both", expand=1, pady=20)
    
    # Creating a Treeview
    tree_ana1 = ttk.Treeview(frame_tab2, columns=list(dafra), show="headings")
    tree_ana1.place(relheight=1, relwidth=1, rely=0)
    
    # Configuring column headings
    for col in dafra:
        tree_ana1.heading(col, text=col,)
        tree_ana1.column(col, width=150, anchor='center')
    
    # Inserting data into the Treeviews
    for index, row in dafra.iterrows():
        values = [row[col] for col in dafra]
        tree_ana1.insert("", "end", values=values)

def graph_tab(tab_, image_path,):
    def on_canvas_scroll(event):
        canvas.yview_scroll(-1 * (event.delta // 120), "units")

    def bind_scroll_event(tab):
        canvas = tab.winfo_children()[0]
        canvas.bind_all("<MouseWheel>", on_canvas_scroll)
        
    # Creating a frame to contain the Graph
    frame_tab = tk.Frame(tab_, )
    frame_tab.place(relheight=0.9, relwidth=1, relx=0, rely=0)

    canvas = tk.Canvas(frame_tab, height=600)
    canvas.place(relheight=1, relwidth=1, relx=0.1, rely=0)

    # Add a vertical scrollbar
    scrollbar = tk.Scrollbar(frame_tab, orient="vertical", command=canvas.yview)
    scrollbar.place(relx=1, rely=0.04, relheight=1, anchor="ne")

    canvas.configure(yscrollcommand=scrollbar.set)

    # Load the saved image and display it within the canvas
    image = Image.open(image_path)
    image = ImageTk.PhotoImage(image)
    label_img = tk.Label(canvas, image=image)
    label_img.image = image
    canvas.create_window((0, 0), window=label_img, anchor="nw")

    canvas.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))
    
    tab_.bind("<Visibility>", lambda event: bind_scroll_event(tab_))
    
    # canvas.bind_all("<MouseWheel>", on_canvas_scroll)

def analysis2_tab(tab_, dafra):
    # Creating a frame to contain the Treeview for tab3
    frame_tab3 = tk.Frame(tab_, height=900)
    frame_tab3.pack(fill="both", expand=1, pady=20)
    
    # Creating a Treeview
    tree_ana2 = ttk.Treeview(frame_tab3, columns=list(dafra), show="headings")
    tree_ana2.place(relheight=1, relwidth=1, rely=0)
    
    # Configuring column headings
    for col in dafra:
        tree_ana2.heading(col, text=col,)
        tree_ana2.column(col, width=120, anchor='center')
    
    # Setting the width of the first 3 columns 
    column_widths = {
        dafra.columns[0]: 60,
        dafra.columns[1]: 90,
        dafra.columns[2]: 110,
        }
    for column, width in column_widths.items():
        tree_ana2.column(column, width=width, anchor="w", stretch=False)
    
    # Inserting data into the Treeviews
    for index, row in dafra.iterrows():
        values = [row[col] for col in dafra]
        tree_ana2.insert("", "end", values=values)

def analysis3_tab(tab_, dafra):
    # Creating a frame to contain the Treeview for tab4
    frame_tab4 = tk.Frame(tab_, height=900)
    frame_tab4.pack(fill="both", expand=1, pady=20)

    # Creating a Treeview
    tree_ana3 = ttk.Treeview(frame_tab4, columns=list(dafra), show="headings")
    tree_ana3.place(relheight=1, relwidth=1, rely=0)
    
    # Configuring column headings
    for col in dafra:
        tree_ana3.heading(col, text=col,)
        tree_ana3.column(col, width=120,)

    # Setting the width of the first 4 columns 
    column_widths = {
        dafra.columns[0]: 60,
        dafra.columns[1]: 90,
        dafra.columns[2]: 110,
        dafra.columns[3]: 100,
        }
    for column, width in column_widths.items():
        tree_ana3.column(column, width=width, anchor="w", stretch=False)

    # Inserting data into the Treeviews
    for index, row in dafra.iterrows():
        values = [row[col] for col in dafra]
        tree_ana3.insert("", "end", values=values)

def control_frame(root, tabview_, data_processor_):
    ########## Creating a control frame for buttons and instructions ##########
    global back_button, next_button
    
    control_frame = tk.Frame(root, )
    control_frame.place(relheight=0.1, relwidth=1, relx=0, rely=0.9)
    control_frame.grid_rowconfigure(0, weight=1)

    # Adding Buttons to control_frame
    back_button = tk.Button(control_frame, image=back_image,
            relief="flat", command= lambda: back_dataframe_button(tabview_), state=tk.DISABLED)
    back_button.place(x=350, y=2, width= 100, height= 55,)

    next_button = tk.Button(control_frame, image=next_image, relief="flat", command= lambda: next_dataframe_button(tabview_))
    next_button.place(x=850, y=2, width= 100, height= 55)

    main_win_button = tk.Button(control_frame, image=main_image, relief='flat', command=main_window_button)
    main_win_button.place(x=100, y=10, width= 130, height= 60)

    export_button = tk.Button(control_frame, image=export_image,relief="flat", command= lambda: export_to_excel(data_processor_))
    export_button.place(x=1160, y=-20, width= 100, height= 100)

def import_file_x():
    global data_processor_X, tabview_x

    # Open a file dialog and load the selected file
    file_path_x = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])

    if file_path_x:
        try:
            destroy_all_widgets(root)

            # Processing the selected file and creating dataframes
            data_processor_X = DataProcessor_X(file_path_x)
            dataframe_result_x = data_processor_X.show_df
            dataframe_percentage_count_df_X = data_processor_X.calculate_percentage_counts()
            percentage_count_df_plot_X = data_processor_X.calculate_percentage_counts_plot()
            dataframe_Subject_percentage_count_df_X = data_processor_X.calculate_subject_percentage_counts()
            Subject_percentage_count_df_plot_X = data_processor_X.calculate_subject_percentage_counts_plot()
            dataframe_highest_marks_df_X = data_processor_X.calculate_highest_marks_students()
            highest_marks_df_plot_X = data_processor_X.calculate_highest_marks_students_plot()

            # Creating a main_frame to contain the Notebook
            main_frame = tk.Frame(root, height=900)
            main_frame.pack(fill="both", expand=1, pady=20)

            # Creating a Tabview
            tabview_x = ttk.Notebook(master=main_frame, height=900)
            tabview_x.pack(fill="both", expand=1, )
            tabview_x.bind("<<NotebookTabChanged>>", on_tab_change_x)

            tab1 = tk.Frame(tabview_x)
            tab2 = tk.Frame(tabview_x)
            tab3 = tk.Frame(tabview_x)
            tab4 = tk.Frame(tabview_x)
            tab5 = tk.Frame(tabview_x)
            tab6 = tk.Frame(tabview_x)
            tab7 = tk.Frame(tabview_x)

            tabview_x.add(tab1, text="Result")
            tabview_x.add(tab2, text="Analysis 1")
            tabview_x.add(tab3, text="Graph 1")
            tabview_x.add(tab4, text="Analysis 2")
            tabview_x.add(tab5, text="Graph 2")
            tabview_x.add(tab6, text="Analysis 3")
            tabview_x.add(tab7, text="Graph 3")

            ############################   TAB 1   ###########################
            result_tab(tab1, dataframe_result_x)
            
            #############################    TAB 2    ###################################
            analysis1_tab(tab2, dataframe_percentage_count_df_X)
            
            #############################    TAB 3    ###################################
            image_path_x1 = 'Items/x_analysis_graph1.png'
            graph_tab(tab3, image_path_x1,)

            #############################    TAB 4    ###################################
            analysis2_tab(tab4, dataframe_Subject_percentage_count_df_X)

            #############################    TAB 5    ###################################
            image_path_x2 = 'Items/x_analysis_graph2.png'
            graph_tab(tab5, image_path_x2,)
            
            #############################    TAB 6    ###################################
            analysis3_tab(tab6, dataframe_highest_marks_df_X)
            
            #############################    TAB 7    ###################################
            image_path_x3 = 'Items/x_analysis_graph3.png'
            graph_tab(tab7, image_path_x3,)
            
            #######   CONTROL FRAME   #########
            control_frame(main_frame, tabview_x, data_processor_X)

        except Exception as e:
            # Show an error message if their is problem while creating the widgets
            messagebox.showerror("Error", "Wrong File Format!")
            main_window_button()
    else:
        # Show an error message if the file format is incorrect
        messagebox.showerror("Error", "Wrong File Extension!")

def import_file_xii():
    global data_processor_XII, tabview_xii

    # Open a file dialog and load the selected file
    file_path_xii = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])

    if file_path_xii:
        try:
            destroy_all_widgets(root)

            # Processing the selected file and creating dataframes
            data_processor_XII = DataProcessor_XII(file_path_xii)
            dataframe_result_xii = data_processor_XII.show_df
            dataframe_result_science = data_processor_XII.science_df
            dataframe_result_pcm = data_processor_XII.pcm_df
            dataframe_result_pcb = data_processor_XII.pcb_df
            dataframe_result_commerce = data_processor_XII.commerce_df
            dataframe_percentage_count_df_Xii = data_processor_XII.calculate_percentage_counts()
            percentage_count_df_plot_XII = data_processor_XII.calculate_percentage_counts_plot()
            dataframe_Subject_percentage_count_df_Xii = data_processor_XII.calculate_subject_percentage_counts()
            Subject_percentage_count_df_plot_XII = data_processor_XII.calculate_subject_percentage_counts_plot()
            dataframe_highest_marks_df_Xii = data_processor_XII.calculate_highest_marks_students()
            highest_marks_df_plot_XII = data_processor_XII.calculate_highest_marks_students_plot()

            # Creating a main_frame to contain the Treeview Frames and scrollbars for tab1
            main_frame = tk.Frame(root, height=900)
            main_frame.pack(fill="both", expand=1, pady=20)
            
            # Creating a Tabview
            tabview_xii = ttk.Notebook(master=main_frame, height=900)
            tabview_xii.pack(fill="both", expand=1, )
            tabview_xii.bind("<<NotebookTabChanged>>", on_tab_change_xii)

            tab1 = tk.Frame(tabview_xii)
            tab2 = tk.Frame(tabview_xii)
            tab3 = tk.Frame(tabview_xii)
            tab4 = tk.Frame(tabview_xii)
            tab5 = tk.Frame(tabview_xii)
            tab6 = tk.Frame(tabview_xii)
            tab7 = tk.Frame(tabview_xii)
            tab8 = tk.Frame(tabview_xii)
            tab9 = tk.Frame(tabview_xii)
            tab10 = tk.Frame(tabview_xii)
            tab11 = tk.Frame(tabview_xii)

            tabview_xii.add(tab1, text="Result")
            tabview_xii.add(tab2, text="Science")
            tabview_xii.add(tab3, text="PCM")
            tabview_xii.add(tab4, text="PCB")
            tabview_xii.add(tab5, text="Commerce")
            tabview_xii.add(tab6, text="Analysis 1")
            tabview_xii.add(tab7, text="Graph 1")
            tabview_xii.add(tab8, text="Analysis 2")
            tabview_xii.add(tab9, text="Graph 2")
            tabview_xii.add(tab10, text="Analysis 3")
            tabview_xii.add(tab11, text="Graph 3")

            ##########################   TAB 1   #########################
            result_tab(tab1, dataframe_result_xii)

            ##########################   TAB 2   #########################
            result_tab(tab2, dataframe_result_science)

            ##########################   TAB 3   #########################
            result_tab(tab3, dataframe_result_pcm)

            ##########################   TAB 4   #########################
            result_tab(tab4, dataframe_result_pcb)

            ##########################   TAB 5   #########################
            result_tab(tab5, dataframe_result_commerce)

            #############################    TAB 6    ###################################
            analysis1_tab(tab6, dataframe_percentage_count_df_Xii)
            
            #############################    TAB 7    ###################################
            image_path_xii1 = 'Items/xii_analysis_graph1.png'
            graph_tab(tab7, image_path_xii1,)
            
            #############################    TAB 8    ###################################
            analysis2_tab(tab8, dataframe_Subject_percentage_count_df_Xii)
            
            #############################    TAB 9    ###################################
            image_path_xii2 = 'Items/xii_analysis_graph2.png'
            graph_tab(tab9, image_path_xii2,)
            
            #############################    TAB 10    ###################################
            analysis3_tab(tab10, dataframe_highest_marks_df_Xii)
            
            #############################    TAB 11    ###################################
            image_path_xii3 = 'Items/xii_analysis_graph3.png'
            graph_tab(tab11, image_path_xii3,)
            
            ##########   CONTROL FRAME   ##########
            control_frame(main_frame, tabview_xii, data_processor_XII)

        except Exception as e:
            # Show an error message if their is problem while creating the widgets
            messagebox.showerror("Error", "Wrong File Format!")
            main_window_button()
    else:
        # Show an error message if the file format is incorrect
        messagebox.showerror("Error", "Wrong File Extension!")

def export_to_excel(df_processor):
    # Function to export the dataframe to Excel
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

    if save_path:
        try:
            # Exporting the dataframe to Excel
            df_processor.save_data_to_excel(save_path)
            df_processor.save_analysis_to_excel(save_path)
            messagebox.showinfo("Success", "Data exported to Excel successfully!")

        except Exception as e:
            messagebox.showerror("Error", str(e))

def back_dataframe_button(tabview_class):
    current_tab = tabview_class.index(tabview_class.select())
    if current_tab > 0:
        tabview_class.select(current_tab - 1)

def next_dataframe_button(tabview_class):
    current_tab = tabview_class.index(tabview_class.select())
    if current_tab < tabview_class.index("end") - 1:
        tabview_class.select(current_tab + 1)

def main_window_button():
    destroy_all_widgets(root)
    my_label=tk.Label(root,image=my_img)
    my_label.place(x=-2, y=-2)
    # Create buttons for selecting class X or XII
    frame1 = tk.Frame(root, relief='sunken', borderwidth=5, background='#cecece')
    frame1.place(x=500, y=200, width=340, height=320)
    text1 = tk.Label(frame1, text="Select Class", font=font.Font(size=24, weight="bold"), borderwidth=3, relief='sunken',
                background='#d9d9d9', foreground='#363737')
    text1.place(x=78, y=50)
    class_x_button = tk.Button(frame1, text=" Class X ", font=font.Font(size=16, weight="normal"), relief="groove",
                command=lambda: select_class("X"), borderwidth=5, background='#d9d9d9')
    class_x_button.place(x=100, y=100, width= 140, height= 70)
    class_xii_button = tk.Button(frame1, text="Class XII", font=font.Font(size=16, weight="normal"), relief="groove",
                command=lambda: select_class("XII"), borderwidth=5, background='#d9d9d9')
    class_xii_button.place(x=100, y=180, width= 140, height= 70)
    
    about_button = tk.Button(frame1, text="❔", relief='groove', command=show_about)
    about_button.place(x=300, y=285,)

def on_tab_change_x(event):
    # Get the currently selected tab
    current_tab = tabview_x.index(tabview_x.select())
    
    # Enable or disable the back and next buttons based on the current tab
    if current_tab == 0:
        back_button.config(state=tk.DISABLED)
        next_button.config(state=tk.NORMAL)
    elif current_tab == 6:
        back_button.config(state=tk.NORMAL)
        next_button.config(state=tk.DISABLED)
    else:
        back_button.config(state=tk.NORMAL)
        next_button.config(state=tk.NORMAL)

def on_tab_change_xii(event):
    # Get the currently selected tab
    current_tab = tabview_xii.index(tabview_xii.select())
    
    # Enable or disable the back and next buttons based on the current tab
    if current_tab == 0:
        back_button.config(state=tk.DISABLED)
        next_button.config(state=tk.NORMAL)
    elif current_tab == 10:
        back_button.config(state=tk.NORMAL)
        next_button.config(state=tk.DISABLED)
    else:
        back_button.config(state=tk.NORMAL)
        next_button.config(state=tk.NORMAL)

def show_about():
    global about_window
    # try:
    #     about_window.destroy()
    # except:
    #     pass
    if about_window is not None:
        about_window.destroy()
    
    about_window = tk.Toplevel(root)
    about_window.title("About")
    about_window.geometry("650x600")
    
    frame = tk.Frame(about_window)
    frame.pack(fill="both", expand=True)
    
    # Create a vertical scrollbar
    vsb = ttk.Scrollbar(frame, orient="vertical")
    vsb.pack(side="right", fill="y")
    
    # Create a horizontal scrollbar
    hsb = ttk.Scrollbar(frame, orient="horizontal")
    hsb.pack(side="bottom", fill="x")
    
    text_widget = Text(frame, wrap="none", yscrollcommand=vsb.set, xscrollcommand=hsb.set, font=font.Font(family='Courier New', size=12, weight='bold'),)
    text_widget.pack(fill="both", expand=True)
    
    vsb.config(command=text_widget.yview)
    hsb.config(command=text_widget.xview)
    
    with open(about, "r") as file:
        about_text = file.read()
        text_widget.insert("1.0", about_text)
        text_widget.config(state="disabled")

# Create buttons for selecting class X or XII
frame1 = tk.Frame(root, relief='sunken', borderwidth=5, background='#cecece')
frame1.place(x=500, y=200, width=340, height=320)
text1 = tk.Label(frame1, text="Select Class", font=font.Font(family='Segoe Print', size=24, weight="bold"), borderwidth=3, relief='sunken',
            background='#d9d9d9', foreground='#363737')
text1.place(x=78, y=50)
class_x_button = tk.Button(frame1, text=" Class X ", font=font.Font(family='Segoe Print', size=16, weight="normal"), relief="groove",
            command=lambda: select_class("X"), borderwidth=5, background='#d9d9d9')
class_x_button.place(x=100, y=130, width= 140, height= 70)
class_xii_button = tk.Button(frame1, text="Class XII", font=font.Font(family='Segoe Print', size=16, weight="normal"), relief="groove",
            command=lambda: select_class("XII"), borderwidth=5, background='#d9d9d9')
class_xii_button.place(x=100, y=210, width= 140, height= 70)

about_window = None  # Initialize about_window as None
about_button = tk.Button(frame1, text="❔", relief='groove', command=show_about)
about_button.place(x=300, y=285,)

# Start the tkinter main loop
root.mainloop()

