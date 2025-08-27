import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time

class SimpleTestGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Simple GUI Test")
        self.root.geometry("600x400")
        
        # Create simple interface
        self.setup_ui()
        
    def setup_ui(self):
        """Set up simple test interface"""
        
        # Title
        title_label = tk.Label(self.root, text="Simple GUI Test", font=('Arial', 16, 'bold'))
        title_label.pack(pady=10)
        
        # Test button
        self.test_button = tk.Button(
            self.root,
            text="Run Test Analysis",
            command=self.run_test,
            font=('Arial', 12),
            bg='#007bff',
            fg='white',
            width=20,
            height=2
        )
        self.test_button.pack(pady=20)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            self.root, 
            mode='indeterminate'
        )
        self.progress_bar.pack(pady=10, padx=20, fill='x')
        
        # Log area
        self.log_text = tk.Text(
            self.root,
            height=15,
            width=70,
            bg='#1e1e1e',
            fg='#00ff00',
            font=('Consolas', 10)
        )
        self.log_text.pack(pady=10, padx=20, fill='both', expand=True)
        
    def log_message(self, message):
        """Add message to log"""
        self.log_text.insert('end', f"{message}\n")
        self.log_text.see('end')
        self.root.update()
        
    def run_test(self):
        """Run a simple test"""
        self.test_button.config(state='disabled')
        self.progress_bar.start()
        
        # Clear log
        self.log_text.delete(1.0, 'end')
        
        # Run test in thread
        thread = threading.Thread(target=self._run_test_thread)
        thread.daemon = True
        thread.start()
        
    def _run_test_thread(self):
        """Run test in background thread"""
        try:
            self.root.after(0, lambda: self.log_message("🌍 Starting test analysis..."))
            time.sleep(2)
            
            self.root.after(0, lambda: self.log_message("📊 Generating sample data..."))
            time.sleep(2)
            
            # Test sample data generation
            from sample_data_generator import generate_sample_esg_data
            df = generate_sample_esg_data()
            
            self.root.after(0, lambda: self.log_message(f"✅ Generated {len(df)} records"))
            time.sleep(1)
            
            # Test data processing
            from esg_analysis_sample import calculate_co2_gdp_ratio, get_latest_year_data, categorize_countries
            
            self.root.after(0, lambda: self.log_message("🔄 Processing data..."))
            df_with_ratios = calculate_co2_gdp_ratio(df)
            latest_data = get_latest_year_data(df_with_ratios)
            final_df, median_ratio = categorize_countries(latest_data)
            
            self.root.after(0, lambda: self.log_message(f"✅ Analysis complete! {len(final_df)} countries processed"))
            self.root.after(0, lambda: self.log_message(f"📈 Median ratio: {median_ratio:.3f}"))
            
            time.sleep(1)
            self.root.after(0, lambda: self.log_message("🎉 Test completed successfully!"))
            
        except Exception as e:
            import traceback
            error_msg = f"❌ Error: {str(e)}"
            traceback_msg = f"❌ Traceback: {traceback.format_exc()}"
            
            self.root.after(0, lambda: self.log_message(error_msg))
            self.root.after(0, lambda: self.log_message(traceback_msg))
            
        finally:
            # Re-enable button and stop progress bar
            self.root.after(0, self._test_complete)
            
    def _test_complete(self):
        """Called when test is complete"""
        self.progress_bar.stop()
        self.test_button.config(state='normal')
        messagebox.showinfo("Test Complete", "GUI test finished! Check the log for details.")

def main():
    """Main function to run the test GUI"""
    root = tk.Tk()
    app = SimpleTestGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()