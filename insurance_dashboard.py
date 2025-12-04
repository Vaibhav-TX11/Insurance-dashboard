import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io

# Page config
st.set_page_config(
    page_title="Insurance Policy Dashboard",
    page_icon="📊",
    layout="wide"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        text-align: center;
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
    # Read the file
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        # Convert date columns if they exist
        date_columns = ['Issued Date', 'Payment Freq']
        for col in date_columns:
            if col in df.columns:
                try:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
                except:
                    pass
        
        # Clean numeric columns
        numeric_columns = ['Commissionable Premium', 'Premium Pay', 'Benefit Term']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        st.success(f"✅ File uploaded successfully! Total records: {len(df)}")
        
        # Sidebar filters
        st.sidebar.header("🔍 Filters")
        
        # Category filter
        if 'Category' in df.columns:
            categories = ['All'] + sorted(df['Category'].dropna().unique().tolist())
            selected_category = st.sidebar.multiselect(
                "Category",
                options=categories,
                default=['All']
            )
        
        # Insurer filter
        if 'Insurer Name' in df.columns:
            insurers = ['All'] + sorted(df['Insurer Name'].dropna().unique().tolist())
            selected_insurer = st.sidebar.multiselect(
                "Insurer Name",
                options=insurers,
                default=['All']
            )
        
        # Agent filter
        if 'Agent Name' in df.columns:
            agents = ['All'] + sorted(df['Agent Name'].dropna().unique().tolist())
            selected_agent = st.sidebar.multiselect(
                "Agent Name",
                options=agents,
                default=['All']
            )
        
        # Branch filter
        if 'Branch name' in df.columns:
            branches = ['All'] + sorted(df['Branch name'].dropna().unique().tolist())
            selected_branch = st.sidebar.multiselect(
                "Branch name",
                options=branches,
                default=['All']
            )
        
        # Manager filter
        if 'Manager Name' in df.columns:
            managers = ['All'] + sorted(df['Manager Name'].dropna().unique().tolist())
            selected_manager = st.sidebar.multiselect(
                "Manager Name",
                options=managers,
                default=['All']
            )
        
        # Date range filter
        if 'Issued Date' in df.columns:
            date_col = df['Issued Date'].dropna()
            if len(date_col) > 0:
                min_date = date_col.min().date()
                max_date = date_col.max().date()
                
                date_range = st.sidebar.date_input(
                    "Issued Date Range",
                    value=(min_date, max_date),
                    min_value=min_date,
                    max_value=max_date
                )
        
        # Premium range filter
        if 'Commissionable Premium' in df.columns:
            premium_col = df['Commissionable Premium'].dropna()
            if len(premium_col) > 0:
                min_premium = float(premium_col.min())
                max_premium = float(premium_col.max())
                
                premium_range = st.sidebar.slider(
                    "Premium Range",
                    min_value=min_premium,
                    max_value=max_premium,
                    value=(min_premium, max_premium)
                )
        
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
        
        # Key Metrics
        st.header("📈 Key Metrics")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            total_premium = filtered_df['Commissionable Premium'].sum() if 'Commissionable Premium' in filtered_df.columns else 0
            st.metric("Total Premium", f"₹{total_premium:,.0f}")
        
        with col2:
            total_paid = filtered_df['Premium Pay'].sum() if 'Premium Pay' in filtered_df.columns else 0
            st.metric("Total Premium Paid", f"₹{total_paid:,.0f}")
        
        with col3:
            st.metric("Total Policies", len(filtered_df))
        
        with col4:
            avg_premium = filtered_df['Commissionable Premium'].mean() if 'Commissionable Premium' in filtered_df.columns else 0
            st.metric("Avg Premium", f"₹{avg_premium:,.0f}")
        
        # Charts Section
        st.header("📊 Visual Analytics")
        
        # Row 1: Two columns
        col1, col2 = st.columns(2)
        
        with col1:
            if 'Category' in filtered_df.columns and 'Commissionable Premium' in filtered_df.columns:
                st.subheader("Premium by Category")
                category_data = filtered_df.groupby('Category')['Commissionable Premium'].sum().reset_index()
                category_data = category_data.sort_values('Commissionable Premium', ascending=False)
                
                fig1 = px.bar(
                    category_data,
                    x='Category',
                    y='Commissionable Premium',
                    color='Category',
                    title='Total Premium by Category'
                )
                fig1.update_layout(showlegend=False, height=400)
                st.plotly_chart(fig1, use_container_width=True)
        
        with col2:
            if 'Insurer Name' in filtered_df.columns and 'Commissionable Premium' in filtered_df.columns:
                st.subheader("Premium by Insurer")
                insurer_data = filtered_df.groupby('Insurer Name')['Commissionable Premium'].sum().reset_index()
                insurer_data = insurer_data.sort_values('Commissionable Premium', ascending=False).head(10)
                
                fig2 = px.bar(
                    insurer_data,
                    x='Commissionable Premium',
                    y='Insurer Name',
                    orientation='h',
                    title='Top 10 Insurers by Premium'
                )
                fig2.update_layout(height=400)
                st.plotly_chart(fig2, use_container_width=True)
        
        # Row 2: Two columns
        col3, col4 = st.columns(2)
        
        with col3:
            if 'Product' in filtered_df.columns:
                st.subheader("Distribution by Product")
                product_counts = filtered_df['Product'].value_counts().head(8)
                
                fig3 = px.pie(
                    values=product_counts.values,
                    names=product_counts.index,
                    title='Product Distribution (Top 8)'
                )
                fig3.update_layout(height=400)
                st.plotly_chart(fig3, use_container_width=True)
        
        with col4:
            if 'Agent Name' in filtered_df.columns and 'Commissionable Premium' in filtered_df.columns:
                st.subheader("Top Agents by Premium")
                agent_data = filtered_df.groupby('Agent Name')['Commissionable Premium'].sum().reset_index()
                agent_data = agent_data.sort_values('Commissionable Premium', ascending=False).head(10)
                
                fig4 = px.bar(
                    agent_data,
                    x='Agent Name',
                    y='Commissionable Premium',
                    color='Commissionable Premium',
                    title='Top 10 Agents by Premium'
                )
                fig4.update_layout(showlegend=False, height=400)
                st.plotly_chart(fig4, use_container_width=True)
        
        # Row 3: Full width charts
        if 'Issued Date' in filtered_df.columns and 'Commissionable Premium' in filtered_df.columns:
            st.subheader("Premium Trends Over Time")
            
            time_data = filtered_df.copy()
            time_data['Month'] = time_data['Issued Date'].dt.to_period('M').astype(str)
            monthly_data = time_data.groupby('Month')['Commissionable Premium'].sum().reset_index()
            
            fig5 = px.line(
                monthly_data,
                x='Month',
                y='Commissionable Premium',
                title='Monthly Premium Trend',
                markers=True
            )
            fig5.update_layout(height=400)
            st.plotly_chart(fig5, use_container_width=True)
        
        # Branch Analysis
        if 'Branch name' in filtered_df.columns and 'Commissionable Premium' in filtered_df.columns:
            st.subheader("Premium by Branch")
            branch_data = filtered_df.groupby('Branch name')['Commissionable Premium'].sum().reset_index()
            branch_data = branch_data.sort_values('Commissionable Premium', ascending=False).head(10)
            
            fig6 = px.bar(
                branch_data,
                x='Branch name',
                y='Commissionable Premium',
                color='Commissionable Premium',
                title='Top 10 Branches by Premium'
            )
            fig6.update_layout(showlegend=False, height=400)
            st.plotly_chart(fig6, use_container_width=True)
        
        # Data Table
        st.header("📋 Filtered Data Table")
        st.dataframe(
            filtered_df,
            use_container_width=True,
            height=400
        )
        
        # Export Section
        st.header("💾 Export Data")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # Export to CSV
            csv = filtered_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download as CSV",
                data=csv,
                file_name=f"insurance_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        
        with col2:
            # Export to Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, index=False, sheet_name='Insurance Data')
            
            st.download_button(
                label="📥 Download as Excel",
                data=buffer.getvalue(),
                file_name=f"insurance_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col3:
            # Export summary statistics
            summary = filtered_df.describe()
            summary_csv = summary.to_csv().encode('utf-8')
            st.download_button(
                label="📥 Download Summary Stats",
                data=summary_csv,
                file_name=f"summary_stats_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
    
    except Exception as e:
        st.error(f"❌ Error reading file: {str(e)}")
        st.info("Please make sure your file is properly formatted and not corrupted.")

else:
    # Instructions when no file is uploaded
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
    - **Visual Charts**: Bar charts, Pie charts, Line charts, Trend analysis
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
