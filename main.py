import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os


class GraphApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel/CSV Graph Plotter")

        self.data_frames = []
        self.annotations = []

        self.initial_dir = os.path.expanduser('~')

        self.button_frame = tk.Frame(root)
        self.button_frame.pack(side=tk.TOP, fill=tk.X)

        self.status_bar = tk.Label(root, text="", bd=1, relief=tk.SUNKEN, anchor=tk.W, font=("Arial", 11))
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.add_graph_label = tk.Label(self.button_frame, text="Add", font=("Arial", 11), cursor="hand2")
        self.add_graph_label.pack(side=tk.LEFT, padx=10)
        self.add_graph_label.bind("<Button-1>", lambda event: self.add_graph())

        self.plot_label = tk.Label(self.button_frame, text="Plot", font=("Arial", 11), cursor="hand2")
        self.plot_label.pack(side=tk.LEFT, padx=10)
        self.plot_label.bind("<Button-1>", lambda event: self.plot_graph())

        self.clear_label = tk.Label(self.button_frame, text="Clear", font=("Arial", 11), cursor="hand2")
        self.clear_label.pack(side=tk.LEFT, padx=10)
        self.clear_label.bind("<Button-1>", lambda event: self.clear_graphs())

        self.reset_label = tk.Label(self.button_frame, text="Reset", font=("Arial", 11), cursor="hand2")
        self.reset_label.pack(side=tk.LEFT, padx=10)
        self.reset_label.bind("<Button-1>", lambda event: self.reset_annotations())

        self.save_image_label = tk.Label(self.button_frame, text="Save", font=("Arial", 11), cursor="hand2")
        self.save_image_label.pack(side=tk.LEFT, padx=10)
        self.save_image_label.bind("<Button-1>", lambda event: self.save_image())

        self.fig, self.ax = plt.subplots(figsize=(12, 6))
        self.canvas = FigureCanvasTkAgg(self.fig, master=root)
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.root.protocol("WM_DELETE_WINDOW", self.close_app)

    def add_graph(self):
        file_path = filedialog.askopenfilename(initialdir=self.initial_dir,
                                               filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")])
        if file_path:
            try:
                df = self.read_file(file_path)
                self.data_frames.append((file_path, df))
                self.update_status(f"Added graph from file: {file_path}")

                self.initial_dir = os.path.dirname(file_path)

            except Exception as e:
                messagebox.showerror("Error", f"Failed to add graph from file: {e}")

    @staticmethod
    def read_file(file_path):
        if file_path.endswith(('.xlsx', '.xls')):
            return pd.read_excel(file_path)
        elif file_path.endswith('.csv'):
            return pd.read_csv(file_path, on_bad_lines='skip')
        else:
            raise ValueError("Unsupported file format")

    def plot_graph(self):
        if not self.data_frames:
            self.update_status("No data to plot. Add graphs first.")
            return

        self.ax.clear()

        num_graphs = len(self.data_frames)
        self.update_status(f"Plotting {num_graphs} graph(s)")

        for file_path, df in self.data_frames:
            for column in df.columns[1:]:
                self.ax.plot(df[df.columns[0]], df[column], label=f'{column} ({os.path.basename(file_path)})')

        self.ax.set_xlabel(self.data_frames[0][1].columns[0])
        self.ax.set_ylabel('Values')
        self.ax.set_title('Graphs from Excel/CSV Data', fontsize=16, fontweight='bold', pad=20)
        self.ax.legend()

        self.fig.canvas.mpl_connect('button_press_event', self.onclick)

        self.canvas.draw()

    def onclick(self, event):
        if event.inaxes:
            x = event.xdata
            y = event.ydata
            annotation = simpledialog.askstring("Add Annotation", "Enter annotation:")
            if annotation:
                self.ax.annotate(annotation, (x, y), textcoords="offset points", xytext=(0, 10), ha='center')
                self.annotations.append((x, y, annotation))
                self.canvas.draw()
                self.update_status(f"Added annotation: '{annotation}'")

    def clear_graphs(self):
        if not self.data_frames:
            self.update_status("No graphs to clear.")
            return

        num_graphs = len(self.data_frames)
        self.data_frames = []
        self.ax.clear()
        self.canvas.draw()
        self.update_status(f"Cleared {num_graphs} graph(s)")

    def reset_annotations(self):
        if not self.annotations:
            self.update_status("No annotations to reset.")
            return

        self.annotations = []
        self.ax.clear()
        self.plot_graph()
        self.update_status("Reset all annotations")

    def save_image(self):
        if not self.data_frames:
            self.update_status("No graph to save. Add and plot graphs first.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".png",
                                                 filetypes=[("PNG files", "*.png"), ("All Files", "*.*")])
        if file_path:
            try:
                self.fig.savefig(file_path)
                self.update_status(f"Saved image: {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save image: {e}")

    def update_status(self, message):
        self.status_bar.config(text=message)

    def close_app(self):
        self.root.quit()


if __name__ == "__main__":
    window = tk.Tk()
    app = GraphApp(window)
    window.mainloop()
