import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import sys
from datetime import datetime
import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Use non-GUI backend for threading safety
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import seaborn as sns

# Import our analysis functions
from sample_data_generator import generate_sample_esg_data
from esg_analysis_sample import (
    calculate_co2_gdp_ratio, 
    get_latest_year_data, 
    categorize_countries,
    create_excel_with_pivot_charts,
    create_visualizations,
    generate_brief
)

class ESGAnalysisGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Rapid ESG Data Insights - GUI")
        self.root.geometry("900x700")
        self.root.configure(bg='#f0f0f0')
        
        # Variables
        self.output_directory = tk.StringVar(value=os.getcwd())
        self.analysis_data = None
        
        self.setup_ui()
        
    def setup_ui(self):
        """Set up the user interface"""
        # Main title
        title_frame = tk.Frame(self.root, bg='#2c3e50', height=60)
        title_frame.pack(fill='x', padx=10, pady=5)
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame, 
            text="🌍 Rapid ESG Data Insights", 
            font=('Arial', 18, 'bold'),
            bg='#2c3e50',
            fg='white'
        )
        title_label.pack(pady=15)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Analysis Tab
        self.analysis_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.analysis_frame, text="📊 Analysis")
        
        # Results Tab
        self.results_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.results_frame, text="📈 Results")
        
        # Charts Tab
        self.charts_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.charts_frame, text="📉 Charts")
        
        self.setup_analysis_tab()
        self.setup_results_tab()
        self.setup_charts_tab()
        
    def setup_analysis_tab(self):
        """Set up the analysis tab"""
        # Configuration Section
        config_frame = ttk.LabelFrame(self.analysis_frame, text="Configuration", padding=10)
        config_frame.pack(fill='x', padx=10, pady=5)
        
        # Output directory selection
        ttk.Label(config_frame, text="Output Directory:").grid(row=0, column=0, sticky='w', pady=5)
        
        dir_frame = ttk.Frame(config_frame)
        dir_frame.grid(row=0, column=1, sticky='ew', pady=5, padx=5)
        config_frame.columnconfigure(1, weight=1)
        
        self.dir_entry = ttk.Entry(dir_frame, textvariable=self.output_directory, width=50)
        self.dir_entry.pack(side='left', fill='x', expand=True)
        
        ttk.Button(
            dir_frame, 
            text="Browse", 
            command=self.browse_directory
        ).pack(side='right', padx=(5, 0))
        
        # Analysis Options
        options_frame = ttk.LabelFrame(self.analysis_frame, text="Analysis Options", padding=10)
        options_frame.pack(fill='x', padx=10, pady=5)
        
        # Data source selection
        ttk.Label(options_frame, text="Data Source:").grid(row=0, column=0, sticky='w', pady=5)
        
        self.data_source = tk.StringVar(value="sample")
        ttk.Radiobutton(
            options_frame, 
            text="Sample Data (Recommended)", 
            variable=self.data_source, 
            value="sample"
        ).grid(row=0, column=1, sticky='w', pady=2)
        
        ttk.Radiobutton(
            options_frame, 
            text="Real World Bank API (2024 Updated Indicators)", 
            variable=self.data_source, 
            value="worldbank"
        ).grid(row=1, column=1, sticky='w', pady=2)
        
        # Generate outputs options
        self.generate_excel = tk.BooleanVar(value=True)
        self.generate_pivot_tables = tk.BooleanVar(value=True)
        self.generate_charts = tk.BooleanVar(value=True)
        self.generate_brief = tk.BooleanVar(value=True)
        
        ttk.Label(options_frame, text="Generate:").grid(row=2, column=0, sticky='nw', pady=5)
        
        outputs_frame = ttk.Frame(options_frame)
        outputs_frame.grid(row=2, column=1, sticky='w', pady=5)
        
        ttk.Checkbutton(outputs_frame, text="Basic Excel Workbook", variable=self.generate_excel).pack(anchor='w')
        ttk.Checkbutton(outputs_frame, text="Enhanced Excel with Pivot Tables", variable=self.generate_pivot_tables).pack(anchor='w')
        ttk.Checkbutton(outputs_frame, text="Visualizations", variable=self.generate_charts).pack(anchor='w')
        ttk.Checkbutton(outputs_frame, text="Executive Brief", variable=self.generate_brief).pack(anchor='w')
        
        # Control Buttons
        control_frame = ttk.Frame(self.analysis_frame)
        control_frame.pack(fill='x', padx=10, pady=10)
        
        self.analyze_button = ttk.Button(
            control_frame,
            text="🚀 Run Analysis",
            command=self.run_analysis,
            style='Accent.TButton'
        )
        self.analyze_button.pack(side='left', padx=5)
        
        ttk.Button(
            control_frame,
            text="📂 Open Output Folder",
            command=self.open_output_folder
        ).pack(side='left', padx=5)
        
        ttk.Button(
            control_frame,
            text="🗑️ Clear Results",
            command=self.clear_results
        ).pack(side='right', padx=5)
        
        # Progress Section
        progress_frame = ttk.LabelFrame(self.analysis_frame, text="Progress", padding=10)
        progress_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            mode='indeterminate'
        )
        self.progress_bar.pack(fill='x', pady=5)
        
        # Log output
        self.log_text = scrolledtext.ScrolledText(
            progress_frame, 
            height=10, 
            state='disabled',
            bg='#1e1e1e',
            fg='#00ff00',
            font=('Consolas', 9)
        )
        self.log_text.pack(fill='both', expand=True, pady=5)
        
    def setup_results_tab(self):
        """Set up the results display tab"""
        # Summary Section
        summary_frame = ttk.LabelFrame(self.results_frame, text="Analysis Summary", padding=10)
        summary_frame.pack(fill='x', padx=10, pady=5)
        
        self.summary_text = scrolledtext.ScrolledText(
            summary_frame, 
            height=8,
            state='disabled',
            font=('Arial', 10)
        )
        self.summary_text.pack(fill='both', expand=True)
        
        # Countries Data Section
        countries_frame = ttk.LabelFrame(self.results_frame, text="Countries Data", padding=10)
        countries_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Treeview for countries data
        columns = ('Country', 'Category', 'CO2/GDP Ratio', 'GDP per Capita', 'CO2 per Capita')
        self.countries_tree = ttk.Treeview(countries_frame, columns=columns, show='headings', height=12)
        
        for col in columns:
            self.countries_tree.heading(col, text=col)
            self.countries_tree.column(col, width=120 if col == 'Country' else 100)
        
        # Scrollbars for treeview
        v_scrollbar = ttk.Scrollbar(countries_frame, orient='vertical', command=self.countries_tree.yview)
        h_scrollbar = ttk.Scrollbar(countries_frame, orient='horizontal', command=self.countries_tree.xview)
        self.countries_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.countries_tree.pack(side='left', fill='both', expand=True)
        v_scrollbar.pack(side='right', fill='y')
        h_scrollbar.pack(side='bottom', fill='x')
        
    def setup_charts_tab(self):
        """Set up the charts display tab"""
        # Chart display area
        self.chart_frame = ttk.Frame(self.charts_frame)
        self.chart_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Chart selection
        chart_control_frame = ttk.Frame(self.charts_frame)
        chart_control_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(chart_control_frame, text="Select Chart:").pack(side='left', padx=5)
        
        self.chart_type = tk.StringVar(value="scatter")
        chart_combo = ttk.Combobox(
            chart_control_frame, 
            textvariable=self.chart_type,
            values=["scatter", "performers", "distribution", "boxplot"],
            state='readonly'
        )
        chart_combo.pack(side='left', padx=5)
        chart_combo.bind('<<ComboboxSelected>>', self.update_chart)
        
        ttk.Button(
            chart_control_frame,
            text="Refresh Chart",
            command=self.update_chart
        ).pack(side='left', padx=5)
        
    def browse_directory(self):
        """Browse for output directory"""
        directory = filedialog.askdirectory(initialdir=self.output_directory.get())
        if directory:
            self.output_directory.set(directory)
            
    def log_message(self, message):
        """Add message to log"""
        self.log_text.config(state='normal')
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert('end', f"[{timestamp}] {message}\n")
        self.log_text.config(state='disabled')
        self.log_text.see('end')
        self.root.update()
        
    def run_analysis(self):
        """Run the ESG analysis in a separate thread"""
        self.analyze_button.config(state='disabled')
        self.progress_bar.start()
        
        # Clear previous results
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, 'end')
        self.log_text.config(state='disabled')
        
        # Run analysis in thread to prevent GUI freezing
        thread = threading.Thread(target=self._run_analysis_thread)
        thread.daemon = True
        thread.start()
        
    def _run_analysis_thread(self):
        """Run analysis in background thread"""
        try:
            # Change to output directory
            original_dir = os.getcwd()
            output_dir = self.output_directory.get()
            self.log_message(f"📁 Working in directory: {output_dir}")
            
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
                self.log_message(f"📁 Created output directory: {output_dir}")
                
            os.chdir(output_dir)
            
            self.log_message("🌍 Starting ESG Data Analysis...")
            
            # Generate or fetch data
            if self.data_source.get() == "sample":
                self.log_message("📊 Generating sample ESG data (30 countries, 2018-2022)...")
                df = generate_sample_esg_data()
            else:
                self.log_message("📊 Fetching real data from World Bank API...")
                from data_processor import WorldBankDataProcessor
                processor = WorldBankDataProcessor()
                
                # Get CO2 data
                self.log_message("🌍 Fetching CO2 emissions data...")
                co2_raw = processor.get_co2_emissions_data()
                if not co2_raw:
                    self.log_message("❌ No CO2 data received from World Bank API")
                    self.log_message("🔄 Falling back to sample data...")
                    df = generate_sample_esg_data()
                else:
                    co2_df = processor.process_data_to_dataframe(co2_raw, 'co2_per_capita')
                    
                    # Get GDP data
                    self.log_message("💰 Fetching GDP per capita data...")
                    gdp_raw = processor.get_gdp_per_capita_data()
                    if not gdp_raw:
                        self.log_message("❌ No GDP data received from World Bank API")
                        self.log_message("🔄 Falling back to sample data...")
                        df = generate_sample_esg_data()
                    else:
                        gdp_df = processor.process_data_to_dataframe(gdp_raw, 'gdp_per_capita')
                        self.log_message(f"✅ Fetched {len(co2_df)} CO2 records and {len(gdp_df)} GDP records")
                        
                        # Combine data
                        combined_df = processor.calculate_co2_gdp_ratio(co2_df, gdp_df)
                        df = processor.get_latest_year_data(combined_df)
                        self.log_message(f"📊 Combined dataset: {len(df)} countries with complete data")
                
            # Process data
            self.log_message("🔄 Processing and calculating CO2/GDP ratios...")
            df_with_ratios = calculate_co2_gdp_ratio(df)
            latest_data = get_latest_year_data(df_with_ratios)
            final_df, median_ratio = categorize_countries(latest_data)
            
            self.analysis_data = final_df
            
            self.log_message(f"✅ Analysis complete! Processed {len(final_df)} countries.")
            self.log_message(f"📈 Median CO2/GDP ratio: {median_ratio:.3f}")
            
            leaders_count = len(final_df[final_df['category'] == 'Leader'])
            laggards_count = len(final_df[final_df['category'] == 'Laggard'])
            
            self.log_message(f"🟢 Leaders: {leaders_count} countries")
            self.log_message(f"🔴 Laggards: {laggards_count} countries")
            
            # Generate outputs based on user selection
            if self.generate_excel.get():
                self.log_message("📊 Creating basic Excel workbook...")
                create_excel_with_pivot_charts(final_df)
                
            if self.generate_pivot_tables.get():
                self.log_message("📊 Creating enhanced Excel workbook with REAL pivot tables...")
                from excel_pivot_enhanced import create_excel_with_real_pivot_tables
                suffix = "_real" if self.data_source.get() == "worldbank" else "_sample"
                create_excel_with_real_pivot_tables(final_df, f'esg_analysis{suffix}_enhanced_pivots.xlsx')
                
            if self.generate_charts.get():
                self.log_message("📈 Generating visualizations...")
                create_visualizations(final_df)
                
            if self.generate_brief.get():
                self.log_message("📝 Creating executive brief...")
                generate_brief(final_df, median_ratio)
                
            self.log_message("🎉 Analysis complete! All outputs generated.")
            
            # Update GUI with results
            self.root.after(0, lambda: self._update_results_display(final_df, median_ratio))
            
        except Exception as e:
            self.log_message(f"❌ Error: {str(e)}")
            self.log_message(f"❌ Error details: {type(e).__name__}")
            import traceback
            self.log_message(f"❌ Traceback: {traceback.format_exc()}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"Analysis failed: {str(e)}"))
        finally:
            # Return to original directory
            try:
                os.chdir(original_dir)
            except:
                pass
            
            # Re-enable button and stop progress bar
            self.root.after(0, self._analysis_complete)
            
    def _analysis_complete(self):
        """Called when analysis is complete"""
        self.progress_bar.stop()
        self.analyze_button.config(state='normal')
        
    def _update_results_display(self, df, median_ratio):
        """Update the results display with analysis data"""
        # Update summary
        leaders = df[df['category'] == 'Leader']
        laggards = df[df['category'] == 'Laggard']
        
        summary = f"""
ANALYSIS SUMMARY
{'='*50}
Total Countries Analyzed: {len(df)}
Median CO2/GDP Ratio: {median_ratio:.3f}

ESG LEADERS: {len(leaders)} countries ({len(leaders)/len(df)*100:.1f}%)
Average CO2/GDP Ratio: {leaders['co2_gdp_ratio'].mean():.3f}

ESG LAGGARDS: {len(laggards)} countries ({len(laggards)/len(df)*100:.1f}%)  
Average CO2/GDP Ratio: {laggards['co2_gdp_ratio'].mean():.3f}

TOP 3 LEADERS:
"""
        top_leaders = leaders.nsmallest(3, 'co2_gdp_ratio')
        for i, (_, row) in enumerate(top_leaders.iterrows(), 1):
            summary += f"{i}. {row['country']} (Ratio: {row['co2_gdp_ratio']:.3f})\n"
            
        summary += "\nTOP 3 LAGGARDS:\n"
        top_laggards = laggards.nlargest(3, 'co2_gdp_ratio')
        for i, (_, row) in enumerate(top_laggards.iterrows(), 1):
            summary += f"{i}. {row['country']} (Ratio: {row['co2_gdp_ratio']:.3f})\n"
        
        self.summary_text.config(state='normal')
        self.summary_text.delete(1.0, 'end')
        self.summary_text.insert(1.0, summary)
        self.summary_text.config(state='disabled')
        
        # Update countries table
        for item in self.countries_tree.get_children():
            self.countries_tree.delete(item)
            
        for _, row in df.iterrows():
            values = (
                row['country'],
                row['category'],
                f"{row['co2_gdp_ratio']:.3f}",
                f"${row['gdp_per_capita']:,.0f}",
                f"{row['co2_per_capita']:.1f}"
            )
            
            tag = 'leader' if row['category'] == 'Leader' else 'laggard'
            self.countries_tree.insert('', 'end', values=values, tags=(tag,))
            
        # Configure row colors
        self.countries_tree.tag_configure('leader', background='#d4edda')
        self.countries_tree.tag_configure('laggard', background='#f8d7da')
        
        # Switch to results tab
        self.notebook.select(self.results_frame)
        
    def update_chart(self, event=None):
        """Update the displayed chart"""
        if self.analysis_data is None:
            messagebox.showwarning("Warning", "Please run analysis first to generate charts.")
            return
            
        try:
            # Clear previous chart
            for widget in self.chart_frame.winfo_children():
                widget.destroy()
                
            # Create new chart based on selection
            plt.ioff()  # Turn off interactive mode
            fig, ax = plt.subplots(figsize=(8, 6))
            
            chart_type = self.chart_type.get()
            df = self.analysis_data
            
            if chart_type == "scatter":
                colors = {'Leader': 'green', 'Laggard': 'red'}
                for category in df['category'].unique():
                    data = df[df['category'] == category]
                    ax.scatter(data['gdp_per_capita'], data['co2_per_capita'], 
                              c=colors[category], label=category, alpha=0.7, s=60)
                
                ax.set_xlabel('GDP per Capita (USD)')
                ax.set_ylabel('CO2 per Capita (metric tons)')
                ax.set_title('CO2 Emissions vs GDP per Capita')
                ax.legend()
                ax.grid(True, alpha=0.3)
                
            elif chart_type == "performers":
                top_5 = df.nlargest(5, 'co2_gdp_ratio')
                bottom_5 = df.nsmallest(5, 'co2_gdp_ratio')
                
                ax.barh(range(len(top_5)), top_5['co2_gdp_ratio'], color='red', alpha=0.7, label='Worst')
                ax.barh(range(len(top_5), len(top_5) + len(bottom_5)), bottom_5['co2_gdp_ratio'], color='green', alpha=0.7, label='Best')
                
                all_countries = list(top_5['country']) + list(bottom_5['country'])
                ax.set_yticks(range(len(all_countries)))
                ax.set_yticklabels(all_countries)
                ax.set_xlabel('CO2/GDP Ratio')
                ax.set_title('Top 5 Best and Worst Performers')
                ax.legend()
                
            elif chart_type == "distribution":
                category_counts = df['category'].value_counts()
                colors = ['green', 'red']
                ax.pie(category_counts.values, labels=category_counts.index, autopct='%1.1f%%', 
                      colors=colors, startangle=90)
                ax.set_title('Distribution of ESG Leaders vs Laggards')
                
            elif chart_type == "boxplot":
                leaders_data = df[df['category'] == 'Leader']['co2_gdp_ratio']
                laggards_data = df[df['category'] == 'Laggard']['co2_gdp_ratio']
                
                box_data = [leaders_data, laggards_data]
                bp = ax.boxplot(box_data, tick_labels=['Leaders', 'Laggards'], patch_artist=True)
                bp['boxes'][0].set_facecolor('green')
                bp['boxes'][0].set_alpha(0.7)
                bp['boxes'][1].set_facecolor('red') 
                bp['boxes'][1].set_alpha(0.7)
                
                ax.set_ylabel('CO2/GDP Ratio')
                ax.set_title('CO2/GDP Ratio Distribution')
                ax.grid(True, alpha=0.3)
        
            plt.tight_layout()
            
            # Embed chart in GUI
            canvas = FigureCanvasTkAgg(fig, self.chart_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)
            
            plt.close(fig)  # Close the figure to free memory
            
        except Exception as e:
            error_label = tk.Label(
                self.chart_frame, 
                text=f"Error creating chart: {str(e)}", 
                fg='red'
            )
            error_label.pack(expand=True)
        
    def open_output_folder(self):
        """Open the output folder in file explorer"""
        import subprocess
        import platform
        
        folder_path = self.output_directory.get()
        
        try:
            if platform.system() == "Darwin":  # macOS
                subprocess.run(["open", folder_path])
            elif platform.system() == "Windows":  # Windows
                subprocess.run(["explorer", folder_path])
            else:  # Linux
                subprocess.run(["xdg-open", folder_path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder: {str(e)}")
            
    def clear_results(self):
        """Clear all results and reset the interface"""
        self.analysis_data = None
        
        # Clear log
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, 'end')
        self.log_text.config(state='disabled')
        
        # Clear summary
        self.summary_text.config(state='normal')
        self.summary_text.delete(1.0, 'end')
        self.summary_text.config(state='disabled')
        
        # Clear countries table
        for item in self.countries_tree.get_children():
            self.countries_tree.delete(item)
            
        # Clear chart
        for widget in self.chart_frame.winfo_children():
            widget.destroy()
            
        messagebox.showinfo("Info", "Results cleared successfully.")

def main():
    """Main function to run the GUI"""
    root = tk.Tk()
    
    # Configure ttk style
    style = ttk.Style()
    style.theme_use('clam')
    
    # Configure accent button style
    style.configure(
        'Accent.TButton',
        background='#007bff',
        foreground='white',
        font=('Arial', 10, 'bold')
    )
    
    app = ESGAnalysisGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()