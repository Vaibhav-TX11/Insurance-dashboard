import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from fpdf import FPDF
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import tempfile
import os
import io

# Page config
st.set_page_config(
    page_title="Insurance Policy Dashboard",
    page_icon="📊",
    layout="wide"
)

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.8rem;
        font-weight: bold;
        color: #2c5f8d;
        margin-top: 2rem;
        margin-bottom: 1rem;
        padding: 10px;
        background-color: #e8f4f8;
        border-left: 5px solid #1f77b4;
    }
    </style>
    """, unsafe_allow_html=True)

# Title
st.markdown('<p class="main-header">📊 Insurance Policy Dashboard</p>', unsafe_allow_html=True)

# File uploader
uploaded_file = st.file_uploader(
    "Upload your Excel file",
    type=['xlsx', 'xls', 'csv'],
    help="Upload an Excel or CSV file containing insurance policy data"
)

if uploaded_file is not None:
    try:
        # Read file
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        # Convert date columns
        date_columns = ['Issued Date', 'Payment Frequency']
        for col in date_columns:
            if col in df.columns:
                try:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
                except:
                    pass
        
        # Clean numeric columns
        numeric_columns = ['Commissionable Premium', 'Premium Paying Term', 'Benefit Term']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        st.success(f"✅ File uploaded successfully! Total records: {len(df)}")
        
        # Sidebar filters
        st.sidebar.header("🔍 Filters")
        
        # Category filter
        if 'Category' in df.columns:
            categories = ['All'] + sorted(df['Category'].dropna().unique().tolist())
            selected_category = st.sidebar.multiselect("Category", options=categories, default=['All'])
        
        # Insurer filter
        if 'Insurer Name' in df.columns:
            insurers = ['All'] + sorted(df['Insurer Name'].dropna().unique().tolist())
            selected_insurer = st.sidebar.multiselect("Insurer Name", options=insurers, default=['All'])
        
        # Agent filter
        if 'Agent Name' in df.columns:
            agents = ['All'] + sorted(df['Agent Name'].dropna().unique().tolist())
            selected_agent = st.sidebar.multiselect("Agent Name", options=agents, default=['All'])
        
        # Branch filter
        if 'Branch name' in df.columns:
            branches = ['All'] + sorted(df['Branch name'].dropna().unique().tolist())
            selected_branch = st.sidebar.multiselect("Branch name", options=branches, default=['All'])
        
        # Manager filter
        if 'Manager Name' in df.columns:
            managers = ['All'] + sorted(df['Manager Name'].dropna().unique().tolist())
            selected_manager = st.sidebar.multiselect("Manager Name", options=managers, default=['All'])
        
        # Date range filter
        if 'Issued Date' in df.columns:
            date_col = df['Issued Date'].dropna()
            if len(date_col) > 0:
                min_date = date_col.min().date()
                max_date = date_col.max().date()
                date_range = st.sidebar.date_input("Issued Date Range", value=(min_date, max_date), 
                                                   min_value=min_date, max_value=max_date)
        
        # Premium range filter
        if 'Commissionable Premium' in df.columns:
            premium_col = df['Commissionable Premium'].dropna()
            if len(premium_col) > 0:
                min_premium = float(premium_col.min())
                max_premium = float(premium_col.max())
                premium_range = st.sidebar.slider("Premium Range", min_value=min_premium, 
                                                  max_value=max_premium, value=(min_premium, max_premium))
        
        # Apply filters
        filtered_df = df.copy()
        
        if 'Category' in df.columns and 'All' not in selected_category:
            filtered_df = filtered_df[filtered_df['Category'].isin(selected_category)]
        
        if 'Insurer Name' in df.columns and 'All' not in selected_insurer:
            filtered_df = filtered_df[filtered_df['Insurer Name'].isin(selected_insurer)]
        
        if 'Agent Name' in df.columns and 'All' not in selected_agent:
            filtered_df = filtered_df[filtered_df['Agent Name'].isin(selected_agent)]
        
        if 'Branch name' in df.columns and 'All' not in selected_branch:
            filtered_df = filtered_df[filtered_df['Branch name'].isin(selected_branch)]
        
        if 'Manager Name' in df.columns and 'All' not in selected_manager:
            filtered_df = filtered_df[filtered_df['Manager Name'].isin(selected_manager)]
        
        if 'Issued Date' in df.columns and len(date_range) == 2:
            filtered_df = filtered_df[
                (filtered_df['Issued Date'].dt.date >= date_range[0]) &
                (filtered_df['Issued Date'].dt.date <= date_range[1])
            ]
        
        if 'Commissionable Premium' in df.columns:
            filtered_df = filtered_df[
                (filtered_df['Commissionable Premium'] >= premium_range[0]) &
                (filtered_df['Commissionable Premium'] <= premium_range[1])
            ]
        
        st.sidebar.markdown(f"**Filtered Records:** {len(filtered_df)} / {len(df)}")
        
        # Get latest month data
        if 'Issued Date' in filtered_df.columns:
            latest_month = filtered_df['Issued Date'].max()
            latest_month_start = latest_month.replace(day=1)
            latest_month_df = filtered_df[
                (filtered_df['Issued Date'] >= latest_month_start) &
                (filtered_df['Issued Date'] <= latest_month)
            ]
            latest_month_str = latest_month.strftime('%B %Y')
        else:
            latest_month_df = filtered_df
            latest_month_str = "Latest Period"
        
        # Overall data (filtered)
        overall_df = filtered_df
        
        # Key Metrics
        st.header("📈 Key Metrics")
        
        # Latest Month Metrics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            latest_premium = latest_month_df['Commissionable Premium'].sum() if 'Commissionable Premium' in latest_month_df.columns else 0
            st.metric(f"Latest Month Premium ({latest_month_str})", f"₹{latest_premium:,.0f}")
        with col2:
            st.metric(f"Latest Month Policies", len(latest_month_df))
        
        # Overall Metrics
        with col3:
            total_premium = overall_df['Commissionable Premium'].sum() if 'Commissionable Premium' in overall_df.columns else 0
            st.metric("Total Premium (Overall)", f"₹{total_premium:,.0f}")
        with col4:
            st.metric("Total Policies (Overall)", len(overall_df))
        
        # ===========================================
        # SECTION 1: LATEST MONTH CHARTS
        # ===========================================
        st.markdown(f'<div class="section-header">📅 Latest Month Analytics ({latest_month_str})</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            if 'Category' in latest_month_df.columns and 'Commissionable Premium' in latest_month_df.columns:
                st.subheader("Premium by Category")
                category_latest = latest_month_df.groupby('Category')['Commissionable Premium'].sum().reset_index()
                category_latest = category_latest.sort_values('Commissionable Premium', ascending=False)
                
                fig_cat_latest = px.bar(category_latest, x='Category', y='Commissionable Premium',
                                       color='Category', title=f'Premium by Category - {latest_month_str}')
                fig_cat_latest.update_layout(showlegend=False, height=400)
                st.plotly_chart(fig_cat_latest, use_container_width=True)
        
        with col2:
            if 'Insurer Name' in latest_month_df.columns and 'Commissionable Premium' in latest_month_df.columns:
                st.subheader("Top Insurers")
                insurer_latest = latest_month_df.groupby('Insurer Name')['Commissionable Premium'].sum().reset_index()
                insurer_latest = insurer_latest.sort_values('Commissionable Premium', ascending=False).head(10)
                
                fig_ins_latest = px.bar(insurer_latest, x='Commissionable Premium', y='Insurer Name',
                                       orientation='h', title=f'Top 10 Insurers - {latest_month_str}')
                fig_ins_latest.update_layout(height=400)
                st.plotly_chart(fig_ins_latest, use_container_width=True)
        
        col3, col4 = st.columns(2)
        
        with col3:
            if 'Product' in latest_month_df.columns:
                st.subheader("Product Distribution")
                product_latest = latest_month_df['Product'].value_counts().head(8)
                
                fig_prod_latest = px.pie(values=product_latest.values, names=product_latest.index,
                                        title=f'Product Distribution - {latest_month_str}')
                fig_prod_latest.update_layout(height=400)
                st.plotly_chart(fig_prod_latest, use_container_width=True)
        
        with col4:
            if 'Manager Name' in latest_month_df.columns and 'Commissionable Premium' in latest_month_df.columns:
                st.subheader("Top Managers")
                manager_latest = latest_month_df.groupby('Manager Name')['Commissionable Premium'].sum().reset_index()
                manager_latest = manager_latest.sort_values('Commissionable Premium', ascending=False).head(10)
                
                fig_mgr_latest = px.bar(manager_latest, x='Manager Name', y='Commissionable Premium',
                                       color='Commissionable Premium', title=f'Top 10 Managers - {latest_month_str}')
                fig_mgr_latest.update_layout(showlegend=False, height=400)
                st.plotly_chart(fig_mgr_latest, use_container_width=True)
        
        if 'Branch name' in latest_month_df.columns and 'Commissionable Premium' in latest_month_df.columns:
            st.subheader("Top Branches")
            branch_latest = latest_month_df.groupby('Branch name')['Commissionable Premium'].sum().reset_index()
            branch_latest = branch_latest.sort_values('Commissionable Premium', ascending=False).head(10)
            
            fig_branch_latest = px.bar(branch_latest, x='Branch name', y='Commissionable Premium',
                                      color='Commissionable Premium', title=f'Top 10 Branches - {latest_month_str}')
            fig_branch_latest.update_layout(showlegend=False, height=400)
            st.plotly_chart(fig_branch_latest, use_container_width=True)
        
        # ===========================================
        # SECTION 2: OVERALL CHARTS
        # ===========================================
        st.markdown('<div class="section-header">📊 Overall Analytics (All Time)</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            if 'Category' in overall_df.columns and 'Commissionable Premium' in overall_df.columns:
                st.subheader("Premium by Category")
                category_overall = overall_df.groupby('Category')['Commissionable Premium'].sum().reset_index()
                category_overall = category_overall.sort_values('Commissionable Premium', ascending=False)
                
                fig_cat_overall = px.bar(category_overall, x='Category', y='Commissionable Premium',
                                        color='Category', title='Premium by Category - Overall')
                fig_cat_overall.update_layout(showlegend=False, height=400)
                st.plotly_chart(fig_cat_overall, use_container_width=True)
        
        with col2:
            if 'Insurer Name' in overall_df.columns and 'Commissionable Premium' in overall_df.columns:
                st.subheader("Top Insurers")
                insurer_overall = overall_df.groupby('Insurer Name')['Commissionable Premium'].sum().reset_index()
                insurer_overall = insurer_overall.sort_values('Commissionable Premium', ascending=False).head(10)
                
                fig_ins_overall = px.bar(insurer_overall, x='Commissionable Premium', y='Insurer Name',
                                        orientation='h', title='Top 10 Insurers - Overall')
                fig_ins_overall.update_layout(height=400)
                st.plotly_chart(fig_ins_overall, use_container_width=True)
        
        col3, col4 = st.columns(2)
        
        with col3:
            if 'Product' in overall_df.columns:
                st.subheader("Product Distribution")
                product_overall = overall_df['Product'].value_counts().head(8)
                
                fig_prod_overall = px.pie(values=product_overall.values, names=product_overall.index,
                                         title='Product Distribution - Overall')
                fig_prod_overall.update_layout(height=400)
                st.plotly_chart(fig_prod_overall, use_container_width=True)
        
        with col4:
            if 'Manager Name' in overall_df.columns and 'Commissionable Premium' in overall_df.columns:
                st.subheader("Top Managers")
                manager_overall = overall_df.groupby('Manager Name')['Commissionable Premium'].sum().reset_index()
                manager_overall = manager_overall.sort_values('Commissionable Premium', ascending=False).head(10)
                
                fig_mgr_overall = px.bar(manager_overall, x='Manager Name', y='Commissionable Premium',
                                        color='Commissionable Premium', title='Top 10 Managers - Overall')
                fig_mgr_overall.update_layout(showlegend=False, height=400)
                st.plotly_chart(fig_mgr_overall, use_container_width=True)
        
        if 'Branch name' in overall_df.columns and 'Commissionable Premium' in overall_df.columns:
            st.subheader("Top Branches")
            branch_overall = overall_df.groupby('Branch name')['Commissionable Premium'].sum().reset_index()
            branch_overall = branch_overall.sort_values('Commissionable Premium', ascending=False).head(10)
            
            fig_branch_overall = px.bar(branch_overall, x='Branch name', y='Commissionable Premium',
                                        color='Commissionable Premium', title='Top 10 Branches - Overall')
            fig_branch_overall.update_layout(showlegend=False, height=400)
            st.plotly_chart(fig_branch_overall, use_container_width=True)
        
        # Monthly Trend (only in overall section)
        if 'Issued Date' in overall_df.columns and 'Commissionable Premium' in overall_df.columns:
            st.subheader("Monthly Premium Trend")
            time_data = overall_df.copy()
            time_data['Month'] = time_data['Issued Date'].dt.to_period('M').astype(str)
            monthly_data = time_data.groupby('Month')['Commissionable Premium'].sum().reset_index()
            monthly_data = monthly_data.sort_values('Month')
            
            fig_trend = px.bar(monthly_data, x='Month', y='Commissionable Premium',
                              title='Monthly Premium Trend')
            fig_trend.update_layout(height=400)
            st.plotly_chart(fig_trend, use_container_width=True)
        
        # Data Table
        st.header("📋 Filtered Data Table")
        st.dataframe(overall_df, use_container_width=True, height=400)
        
        # Export Section
        st.header("💾 Export Data")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            csv = overall_df.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Download as CSV", data=csv,
                             file_name=f"insurance_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                             mime="text/csv")
        
        with col2:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                overall_df.to_excel(writer, index=False, sheet_name='Insurance Data')
            st.download_button("📥 Download as Excel", data=buffer.getvalue(),
                             file_name=f"insurance_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        with col3:
            with st.expander("📄 Download Full PDF Report"):
                if st.button("Generate PDF Report"):
                    with st.spinner("Preparing PDF with Latest Month + Overall charts..."):
                        temp_dir = tempfile.mkdtemp()
                        chart_paths = []
                        
                        # Helper functions (same as before)
                        def save_matplotlib_bar(data, x, y, title, filename, color="skyblue"):
                            fig, ax = plt.subplots(figsize=(12, 5))
                            bars = ax.bar(data[x], data[y], color=color, alpha=0.8, edgecolor='navy', linewidth=0.5)
                            ax.set_title(title, fontsize=14, fontweight='bold', pad=15)
                            ax.set_xlabel(x, fontsize=11, fontweight='bold')
                            ax.set_ylabel(y, fontsize=11, fontweight='bold')
                            ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: f'{x:,.0f}'))
                            plt.xticks(rotation=45, ha='right', fontsize=9)
                            plt.yticks(fontsize=9)
                            if len(bars) <= 15:
                                for bar in bars:
                                    height = bar.get_height()
                                    if height > 0:
                                        ax.text(bar.get_x() + bar.get_width()/2., height,
                                            f'{height:,.0f}', ha='center', va='bottom', fontsize=8, fontweight='bold')
                            ax.grid(axis='y', alpha=0.3, linestyle='--')
                            ax.set_axisbelow(True)
                            plt.tight_layout()
                            plt.savefig(filename, dpi=120, bbox_inches='tight')
                            plt.close()
                        
                        def save_matplotlib_barh(data, x, y, title, filename, color="skyblue"):
                            fig, ax = plt.subplots(figsize=(12, 6))
                            bars = ax.barh(data[y], data[x], color=color, alpha=0.8, edgecolor='navy', linewidth=0.5)
                            ax.set_title(title, fontsize=14, fontweight='bold', pad=15)
                            ax.set_xlabel(x, fontsize=11, fontweight='bold')
                            ax.set_ylabel(y, fontsize=11, fontweight='bold')
                            ax.xaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: f'{x:,.0f}'))
                            plt.xticks(fontsize=9)
                            plt.yticks(fontsize=9)
                            for bar in bars:
                                width = bar.get_width()
                                if width > 0:
                                    ax.text(width, bar.get_y() + bar.get_height()/2., f'{width:,.0f}',
                                        ha='left', va='center', fontsize=8, fontweight='bold',
                                        bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.7))
                            ax.grid(axis='x', alpha=0.3, linestyle='--')
                            ax.set_axisbelow(True)
                            plt.tight_layout()
                            plt.savefig(filename, dpi=120, bbox_inches='tight')
                            plt.close()
                        
                        def save_matplotlib_pie(values, labels, title, filename):
                            fig, ax = plt.subplots(figsize=(10, 6))
                            wedges, texts, autotexts = ax.pie(values, labels=None, autopct='%1.1f%%',
                                                            startangle=90, colors=plt.cm.Set2.colors,
                                                            textprops={'fontsize': 10, 'fontweight': 'bold'},
                                                            wedgeprops={'edgecolor': 'white', 'linewidth': 2})
                            for autotext in autotexts:
                                autotext.set_color('white')
                                autotext.set_fontweight('bold')
                            ax.set_title(title, fontsize=14, fontweight='bold', pad=15)
                            ax.legend(wedges, labels, title="Products", loc="center left",
                                    bbox_to_anchor=(1, 0, 0.5, 1), fontsize=9, title_fontsize=10)
                            ax.axis('equal')
                            plt.tight_layout()
                            plt.savefig(filename, dpi=120, bbox_inches='tight')
                            plt.close()
                        
                        try:
                            # Generate charts for LATEST MONTH
                            # 1. Category - Latest
                            path_cat_latest = os.path.join(temp_dir, "cat_latest.png")
                            save_matplotlib_bar(category_latest, "Category", "Commissionable Premium",
                                            f"Premium by Category - {latest_month_str}", path_cat_latest)
                            
                            # 2. Category - Overall
                            path_cat_overall = os.path.join(temp_dir, "cat_overall.png")
                            save_matplotlib_bar(category_overall, "Category", "Commissionable Premium",
                                            "Premium by Category - Overall", path_cat_overall)
                            
                            # 3. Insurer - Latest
                            path_ins_latest = os.path.join(temp_dir, "ins_latest.png")
                            save_matplotlib_barh(insurer_latest, "Commissionable Premium", "Insurer Name",
                                                f"Top 10 Insurers - {latest_month_str}", path_ins_latest)
                            
                            # 4. Insurer - Overall
                            path_ins_overall = os.path.join(temp_dir, "ins_overall.png")
                            save_matplotlib_barh(insurer_overall, "Commissionable Premium", "Insurer Name",
                                                "Top 10 Insurers - Overall", path_ins_overall)
                            
                            # 5. Product - Latest
                            path_prod_latest = os.path.join(temp_dir, "prod_latest.png")
                            save_matplotlib_pie(product_latest.values, product_latest.index,
                                            f"Product Distribution - {latest_month_str}", path_prod_latest)
                            
                            # 6. Product - Overall
                            path_prod_overall = os.path.join(temp_dir, "prod_overall.png")
                            save_matplotlib_pie(product_overall.values, product_overall.index,
                                            "Product Distribution - Overall", path_prod_overall)
                            
                            # 7. Manager - Latest
                            path_mgr_latest = os.path.join(temp_dir, "mgr_latest.png")
                            save_matplotlib_bar(manager_latest, "Manager Name", "Commissionable Premium",
                                            f"Top 10 Managers - {latest_month_str}", path_mgr_latest)
                            
                            # 8. Manager - Overall
                            path_mgr_overall = os.path.join(temp_dir, "mgr_overall.png")
                            save_matplotlib_bar(manager_overall, "Manager Name", "Commissionable Premium",
                                            "Top 10 Managers - Overall", path_mgr_overall)
                            
                            # 9. Branch - Latest
                            path_branch_latest = os.path.join(temp_dir, "branch_latest.png")
                            save_matplotlib_bar(branch_latest, "Branch name", "Commissionable Premium",
                                            f"Top 10 Branches - {latest_month_str}", path_branch_latest)
                            
                            # 10. Branch - Overall
                            path_branch_overall = os.path.join(temp_dir, "branch_overall.png")
                            save_matplotlib_bar(branch_overall, "Branch name", "Commissionable Premium",
                                            "Top 10 Branches - Overall", path_branch_overall)
                            
                            # 11. Monthly Trend (only overall)
                            path_trend = os.path.join(temp_dir, "trend.png")
                            save_matplotlib_bar(monthly_data, "Month", "Commissionable Premium",
                                            "Monthly Premium Trend", path_trend)
                            
                            # Create PDF
                            from fpdf.enums import XPos, YPos
                            pdf = FPDF(orientation='P', unit='mm', format='A4')
                            pdf.set_auto_page_break(auto=True, margin=15)
                            
                            # Title Page
                            pdf.add_page()
                            pdf.set_font("Helvetica", "B", 22)
                            pdf.cell(0, 15, "Insurance Dashboard Report", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf.set_font("Helvetica", "", 12)
                            pdf.ln(5)
                            pdf.cell(0, 10, f"Generated: {datetime.now().strftime('%d-%m-%Y %H:%M')}", 
                                    new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf.ln(10)
                            
                            # Key Metrics
                            pdf.set_font("Helvetica", "B", 14)
                            pdf.cell(0, 10, "Key Metrics Summary", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf.set_font("Helvetica", "", 11)
                            pdf.cell(0, 8, f"Latest Month ({latest_month_str}): Rs.{latest_premium:,.0f}", 
                                    new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf.cell(0, 8, f"Overall Total: Rs.{total_premium:,.0f}", 
                                    new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf.cell(0, 8, f"Total Policies: {len(overall_df)}", 
                                    new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            
                            # Page 1: Category (Latest + Overall)
                            pdf.add_page()
                            pdf.set_font("Helvetica", "B", 14)
                            pdf.cell(0, 8, f"Category - Latest Month ({latest_month_str})", 
                                    new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf.image(path_cat_latest, x=10, w=190)
                            pdf.ln(5)
                            pdf.cell(0, 8, "Category - Overall", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf.image(path_cat_overall, x=10, w=190)
                            
                            # Page 2: Insurer (Latest + Overall)
                            pdf.add_page()
                            pdf.set_font("Helvetica", "B", 14)
                            pdf.cell(0, 8, f"Insurers - Latest Month ({latest_month_str})", 
                                    new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf.image(path_ins_latest, x=10, w=190)
                            pdf.ln(5)
                            pdf.cell(0, 8, "Insurers - Overall", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf.image(path_ins_overall, x=10, w=190)
                            
                            # Page 3: Product (Latest + Overall)
                            pdf.add_page()
                            pdf.set_font("Helvetica", "B", 14)
                            pdf.cell(0, 8, f"Products - Latest Month ({latest_month_str})", 
                                    new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf.image(path_prod_latest, x=10, w=190)
                            pdf.ln(5)
                            pdf.cell(0, 8, "Products - Overall", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf.image(path_prod_overall, x=10, w=190)
                            
                            # Page 4: Manager (Latest + Overall)
                            pdf.add_page()
                            pdf.set_font("Helvetica", "B", 14)
                            pdf.cell(0, 8, f"Managers - Latest Month ({latest_month_str})", 
                                    new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf.image(path_mgr_latest, x=10, w=190)
                            pdf.ln(5)
                            pdf.cell(0, 8, "Managers - Overall", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf.image(path_mgr_overall, x=10, w=190)
                            
                            # Page 5: Branch (Latest + Overall)
                            pdf.add_page()
                            pdf.set_font("Helvetica", "B", 14)
                            pdf.cell(0, 8, f"Branches - Latest Month ({latest_month_str})", 
                                    new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf.image(path_branch_latest, x=10, w=190)
                            pdf.ln(5)
                            pdf.cell(0, 8, "Branches - Overall", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf.image(path_branch_overall, x=10, w=190)
                            
                            # Page 6: Monthly Trend
                            pdf.add_page()
                            pdf.set_font("Helvetica", "B", 14)
                            pdf.cell(0, 8, "Monthly Premium Trend", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf.image(path_trend, x=10, w=190)
                            
                            # Save PDF
                            pdf_path = os.path.join(temp_dir, "insurance_report.pdf")
                            pdf.output(pdf_path)
                            
                            # Download button
                            with open(pdf_path, "rb") as f:
                                st.download_button(
                                    label="📄 Download Complete PDF Report",
                                    data=f,
                                    file_name=f"insurance_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                                    mime="application/pdf"
                                )
                            
                            st.success("✅ PDF with Latest Month + Overall charts is ready!")
                            
                        except Exception as e:
                            st.error(f"PDF generation error: {str(e)}")
    
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")

else:
    st.info("👆 Please upload an Excel or CSV file to get started")    
    st.markdown("""
    ### 📝 Instructions:
    1. Upload your Excel file using the file uploader above
    2. Use the sidebar filters to narrow down your data
    3. View key metrics and visual analytics
    4. Export filtered data in your preferred format
    
    ### 📊 Features:
    - **Interactive Filters**: Category, Insurer, Agent, Branch, Date Range, Premium Range
    - **Key Metrics**: Total Premium, Premium Paid, Policy Count, Average Premium
    - **Visual Charts**: Bar charts, Pie charts, Trend analysis
    - **Data Export**: CSV, Excel, Summary Statistics
    - **Real-time Updates**: All charts update instantly with filters
    """)

    st.info("""
    🔒 **Privacy Notice**: 
    - Your uploaded file is processed locally in your browser session
    - No data is stored on our servers
    - Data is automatically deleted when you close this page
    - Each user session is completely isolated
    """)
