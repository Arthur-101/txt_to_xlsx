from customtkinter import *

# Main Window
def main():
    
    set_appearance_mode('dark')
    set_default_color_theme('green')

    # root : Creating a master
    width = 920
    height = 615
    root = CTk()
    root.title('Banking System')
    root.geometry(f"{width}x{height}")
    root.geometry("+170+50")
    
    
    root.mainloop()

main()
