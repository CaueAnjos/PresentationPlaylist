import tkinter as tk
from tkinter import filedialog, messagebox
import os


class PlaylistApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PowerPoint Playlist Manager")

        self.presentation_list = []

        # Listbox to show the selected presentations
        self.listbox = tk.Listbox(root, width=50, height=10)
        self.listbox.pack(pady=20)

        # Buttons to add, remove, move up, move down
        self.add_button = tk.Button(
            root, text="Add Presentations", command=self.add_presentations
        )
        self.add_button.pack(pady=5)

        self.remove_button = tk.Button(
            root, text="Remove Selected", command=self.remove_selected
        )
        self.remove_button.pack(pady=5)

        self.move_up_button = tk.Button(root, text="Move Up", command=self.move_up)
        self.move_up_button.pack(pady=5)

        self.move_down_button = tk.Button(
            root, text="Move Down", command=self.move_down
        )
        self.move_down_button.pack(pady=5)

        self.save_button = tk.Button(
            root, text="Save Playlist", command=self.save_playlist
        )
        self.save_button.pack(pady=20)

    def add_presentations(self):
        files = filedialog.askopenfilenames(
            title="Select PowerPoint Presentations",
            filetypes=[("PowerPoint files", "*.pptx")],
        )
        for file in files:
            self.presentation_list.append(file)
            self.listbox.insert(tk.END, os.path.basename(file))

    def remove_selected(self):
        selected = self.listbox.curselection()
        if selected:
            for index in selected[::-1]:
                self.listbox.delete(index)
                del self.presentation_list[index]

    def move_up(self):
        selected = self.listbox.curselection()
        if selected:
            for index in selected:
                if index > 0:
                    # Swap in listbox
                    self.listbox.insert(index - 1, self.listbox.get(index))
                    self.listbox.delete(index + 1)
                    # Swap in presentation list
                    self.presentation_list[index - 1], self.presentation_list[index] = (
                        self.presentation_list[index],
                        self.presentation_list[index - 1],
                    )

    def move_down(self):
        selected = self.listbox.curselection()
        if selected:
            for index in selected:
                if index < len(self.presentation_list) - 1:
                    # Swap in listbox
                    self.listbox.insert(index + 2, self.listbox.get(index))
                    self.listbox.delete(index)
                    # Swap in presentation list
                    self.presentation_list[index + 1], self.presentation_list[index] = (
                        self.presentation_list[index],
                        self.presentation_list[index + 1],
                    )

    def save_playlist(self):
        if not self.presentation_list:
            messagebox.showerror("Error", "No presentations selected.")
            return

        # Save the playlist as a text file
        playlist_path = filedialog.asksaveasfilename(
            title="Save Playlist",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")],
        )
        if playlist_path:
            with open(playlist_path, "w") as file:
                for presentation in self.presentation_list:
                    file.write(f"{presentation}\n")
            messagebox.showinfo("Success", "Playlist saved successfully.")
