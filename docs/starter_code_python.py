"""
PROJECT 2: E-Commerce Business Analytics
Group Project - Starter Code (Python)
Saturday: 9.5 hours work time

TEAM INFORMATION:
Team Name: _______________________
Members:
1. _________________ (Role: _____________)
2. _________________ (Role: _____________)
3. _________________ (Role: _____________)
4. _________________ (Role: _____________)
5. _________________ (Role: _____________)

DELIVERABLES CHECKLIST:
[ ] Dashboard (HTML/Notebook)
[ ] PowerPoint Presentation (10-15 slides)
[ ] Technical Report (4-6 pages)
[ ] Clean code with README
[ ] All 10 analyses complete
"""

# ==============================================================================
# SETUP & IMPORTS
# ==============================================================================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Set styles
sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (14, 8)

# Professional color palette
COLORS = {
    'primary': '#1f77b4',
    'secondary': '#ff7f0e',
    'success': '#2ca02c',
    'danger': '#d62728',
    'info': '#17becf'
}

print("="*70)
print(" "*20 + "E-COMMERCE ANALYTICS PROJECT")
print("="*70)
print(f"Team: [YOUR TEAM NAME]")
print(f"Start Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print("="*70)

# ==============================================================================
# DATA LOADING
# ==============================================================================

print("\n📊 LOADING DATA...")

# Load the main dataset
# NOTE: Update the file path to where you saved the data
file_path = "online_retail.csv"  # UPDATE THIS PATH

try:
    df = pd.read_csv(file_path, encoding='ISO-8859-1')
    print(f"✓ Data loaded successfully: {df.shape[0]:,} rows × {df.shape[1]} columns")
except FileNotFoundError:
    print("❌ Error: Data file not found!")
    print("Please download from: https://archive.ics.uci.edu/ml/datasets/Online+Retail")
    print("Or check your file path")

# Display basic info
print("\nColumn Names:")
print(df.columns.tolist())

print("\nFirst 5 rows:")
print(df.head())

print("\nData Types:")
print(df.dtypes)

# ==============================================================================
# DATA EXPLORATION
# ==============================================================================

print("\n"  + "="*70)
print(" "*25 + "DATA EXPLORATION")
print("="*70)

print("\n📈 BASIC STATISTICS:")
print(f"Total Transactions: {df['InvoiceNo'].nunique():,}")
print(f"Unique Products: {df['StockCode'].nunique():,}")
print(f"Unique Customers: {df['CustomerID'].nunique():,}")
print(f"Countries: {df['Country'].nunique():,}")
print(f"Date Range: {df['InvoiceDate'].min()} to {df['InvoiceDate'].max()}")

print("\n❓ MISSING VALUES:")
print(df.isnull().sum())

print("\n💰 PRICE STATISTICS:")
print(df['UnitPrice'].describe())

print("\n📦 QUANTITY STATISTICS:")
print(df['Quantity'].describe())

# ==============================================================================
# DATA CLEANING
# ==============================================================================

print("\n" + "="*70)
print(" "*25 + "DATA CLEANING")
print("="*70)

# Create a copy for cleaning
df_clean = df.copy()

# 1. Convert InvoiceDate to datetime
df_clean['InvoiceDate'] = pd.to_datetime(df_clean['InvoiceDate'])
print("✓ Converted InvoiceDate to datetime")

# 2. Create calculated fields
df_clean['TotalPrice'] = df_clean['Quantity'] * df_clean['UnitPrice']
print("✓ Created TotalPrice column")

# 3. Extract time components
df_clean['Year'] = df_clean['InvoiceDate'].dt.year
df_clean['Month'] = df_clean['InvoiceDate'].dt.month
df_clean['MonthName'] = df_clean['InvoiceDate'].dt.month_name()
df_clean['Day'] = df_clean['InvoiceDate'].dt.day
df_clean['DayOfWeek'] = df_clean['InvoiceDate'].dt.day_name()
df_clean['Hour'] = df_clean['InvoiceDate'].dt.hour
df_clean['Date'] = df_clean['InvoiceDate'].dt.date
print("✓ Extracted time components")

# 4. TODO: Handle missing CustomerID
# Decision: How do you want to handle transactions without CustomerID?
# Option A: Remove them
# Option B: Keep them but exclude from customer-specific analyses
# Option C: Create a "Guest" category

# YOUR DECISION HERE:



# 5. TODO: Handle negative quantities (returns/cancellations)
# Decision: How do you want to handle returns?
# Option A: Remove all negative quantities
# Option B: Create separate returns dataset
# Option C: Keep them for return analysis

# YOUR DECISION HERE:



# 6. TODO: Handle outliers in price and quantity
# Check for unusual values and decide how to handle

# YOUR CODE HERE:



# 7. Remove invalid prices (if any)
df_clean = df_clean[df_clean['UnitPrice'] > 0]
print(f"✓ Removed {len(df) - len(df_clean):,} rows with invalid prices")

print(f"\n✓ CLEANED DATA: {df_clean.shape[0]:,} rows × {df_clean.shape[1]} columns")

# ==============================================================================
# ANALYSIS 1: SALES OVERVIEW & TRENDS
# ==============================================================================

print("\n" + "="*70)
print(" "*20 + "ANALYSIS 1: SALES OVERVIEW & TRENDS")
print("="*70)

# Calculate key metrics
total_revenue = df_clean['TotalPrice'].sum()
total_transactions = df_clean['InvoiceNo'].nunique()
total_customers = df_clean['CustomerID'].nunique()
avg_order_value = df_clean.groupby('InvoiceNo')['TotalPrice'].sum().mean()

print(f"\n💰 TOTAL REVENUE: ${total_revenue:,.2f}")
print(f"🛒 TOTAL TRANSACTIONS: {total_transactions:,}")
print(f"👥 TOTAL CUSTOMERS: {total_customers:,}")
print(f"💵 AVERAGE ORDER VALUE: ${avg_order_value:,.2f}")

# Monthly revenue trend
monthly_revenue = df_clean.groupby(df_clean['InvoiceDate'].dt.to_period('M'))['TotalPrice'].sum()
monthly_revenue.index = monthly_revenue.index.to_timestamp()

fig1 = px.line(
    x=monthly_revenue.index,
    y=monthly_revenue.values,
    title='Monthly Revenue Trend',
    labels={'x': 'Month', 'y': 'Revenue ($)'},
    markers=True
)
fig1.update_layout(height=500)
fig1.show()
fig1.write_html("analysis1_monthly_revenue.html")

# TODO: Add more trend visualizations
# - Weekly trends
# - Growth rate calculation
# - Seasonal patterns



print("✓ Analysis 1 complete")

# ==============================================================================
# ANALYSIS 2: PRODUCT PERFORMANCE
# ==============================================================================

print("\n" + "="*70)
print(" "*20 + "ANALYSIS 2: PRODUCT PERFORMANCE")
print("="*70)

# Top 10 products by quantity
top_products_qty = df_clean.groupby('Description')['Quantity'].sum().nlargest(10)

fig2a = px.bar(
    x=top_products_qty.values,
    y=top_products_qty.index,
    orientation='h',
    title='Top 10 Best-Selling Products (by Quantity)',
    labels={'x': 'Quantity Sold', 'y': 'Product'},
    color=top_products_qty.values,
    color_continuous_scale='Blues'
)
fig2a.update_layout(height=500, showlegend=False)
fig2a.show()
fig2a.write_html("analysis2_top_products_quantity.html")

# Top 10 products by revenue
top_products_rev = df_clean.groupby('Description')['TotalPrice'].sum().nlargest(10)

fig2b = px.bar(
    x=top_products_rev.values,
    y=top_products_rev.index,
    orientation='h',
    title='Top 10 Products by Revenue',
    labels={'x': 'Revenue ($)', 'y': 'Product'},
    color=top_products_rev.values,
    color_continuous_scale='Greens'
)
fig2b.update_layout(height=500, showlegend=False)
fig2b.show()
fig2b.write_html("analysis2_top_products_revenue.html")

# TODO: Additional product analyses
# - Price vs sales volume
# - Underperforming products
# - Product category analysis (if you create categories)



print("✓ Analysis 2 complete")

# ==============================================================================
# ANALYSIS 3: CUSTOMER BEHAVIOR
# ==============================================================================

print("\n" + "="*70)
print(" "*20 + "ANALYSIS 3: CUSTOMER BEHAVIOR")
print("="*70)

# Customer purchase frequency
customer_freq = df_clean.groupby('CustomerID')['InvoiceNo'].nunique()

print(f"\n📊 CUSTOMER PURCHASE FREQUENCY:")
print(f"One-time buyers: {(customer_freq == 1).sum():,} ({(customer_freq == 1).sum()/len(customer_freq)*100:.1f}%)")
print(f"Repeat customers: {(customer_freq > 1).sum():,} ({(customer_freq > 1).sum()/len(customer_freq)*100:.1f}%)")
print(f"Average purchases per customer: {customer_freq.mean():.2f}")

# Distribution of purchase frequency
fig3a = px.histogram(
    customer_freq,
    nbins=50,
    title='Distribution of Customer Purchase Frequency',
    labels={'value': 'Number of Purchases', 'count': 'Number of Customers'}
)
fig3a.update_layout(height=500)
fig3a.show()
fig3a.write_html("analysis3_purchase_frequency.html")

# Customer Lifetime Value (CLV)
customer_clv = df_clean.groupby('CustomerID')['TotalPrice'].sum().sort_values(ascending=False)

print(f"\n💰 TOP 10 CUSTOMERS BY CLV:")
print(customer_clv.head(10))

# TODO: Additional customer analyses
# - Customer segmentation (RFM analysis)
# - New vs returning customer trends
# - Customer retention analysis



print("✓ Analysis 3 complete")

# ==============================================================================
# ANALYSIS 4: GEOGRAPHIC ANALYSIS
# ==============================================================================

print("\n" + "="*70)
print(" "*20 + "ANALYSIS 4: GEOGRAPHIC ANALYSIS")
print("="*70)

# Revenue by country
country_revenue = df_clean.groupby('Country')['TotalPrice'].sum().sort_values(ascending=False)

print(f"\n🌍 TOP 10 COUNTRIES BY REVENUE:")
for i, (country, revenue) in enumerate(country_revenue.head(10).items(), 1):
    print(f"{i}. {country}: ${revenue:,.2f}")

# Visualize top countries
fig4a = px.bar(
    x=country_revenue.head(15).values,
    y=country_revenue.head(15).index,
    orientation='h',
    title='Top 15 Countries by Revenue',
    labels={'x': 'Revenue ($)', 'y': 'Country'},
    color=country_revenue.head(15).values,
    color_continuous_scale='Viridis'
)
fig4a.update_layout(height=600, showlegend=False)
fig4a.show()
fig4a.write_html("analysis4_country_revenue.html")

# TODO: Create geographic map visualization
# Hint: Use plotly choropleth or scatter_geo



print("✓ Analysis 4 complete")

# ==============================================================================
# ANALYSIS 5: TIME-BASED PATTERNS
# ==============================================================================

print("\n" + "="*70)
print(" "*20 + "ANALYSIS 5: TIME-BASED PATTERNS")
print("="*70)

# Sales by day of week
dow_revenue = df_clean.groupby('DayOfWeek')['TotalPrice'].sum()
dow_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
dow_revenue = dow_revenue.reindex(dow_order)

fig5a = px.bar(
    x=dow_order,
    y=dow_revenue.values,
    title='Revenue by Day of Week',
    labels={'x': 'Day of Week', 'y': 'Revenue ($)'},
    color=dow_revenue.values,
    color_continuous_scale='Blues'
)
fig5a.update_layout(height=500, showlegend=False)
fig5a.show()
fig5a.write_html("analysis5_day_of_week.html")

# Sales by hour of day
hour_revenue = df_clean.groupby('Hour')['TotalPrice'].sum()

fig5b = px.line(
    x=hour_revenue.index,
    y=hour_revenue.values,
    title='Revenue by Hour of Day',
    labels={'x': 'Hour', 'y': 'Revenue ($)'},
    markers=True
)
fig5b.update_layout(height=500)
fig5b.show()
fig5b.write_html("analysis5_hour_of_day.html")

# TODO: Additional time analyses
# - Monthly seasonality
# - Weekend vs weekday
# - Holiday effects (if you identify holidays)



print("✓ Analysis 5 complete")

# ==============================================================================
# ANALYSIS 6: BASKET ANALYSIS
# ==============================================================================

print("\n" + "="*70)
print(" "*20 + "ANALYSIS 6: BASKET ANALYSIS")
print("="*70)

# Calculate basket metrics
basket_size = df_clean.groupby('InvoiceNo')['Quantity'].sum()
basket_value = df_clean.groupby('InvoiceNo')['TotalPrice'].sum()
basket_items = df_clean.groupby('InvoiceNo')['StockCode'].count()

print(f"\n🛒 BASKET STATISTICS:")
print(f"Average items per basket: {basket_items.mean():.2f}")
print(f"Average basket value: ${basket_value.mean():.2f}")
print(f"Median basket value: ${basket_value.median():.2f}")

# Distribution of basket values
fig6 = px.histogram(
    basket_value[basket_value < basket_value.quantile(0.95)],  # Remove top 5% outliers for visibility
    nbins=50,
    title='Distribution of Basket Values (excluding top 5% outliers)',
    labels={'value': 'Basket Value ($)', 'count': 'Frequency'}
)
fig6.update_layout(height=500)
fig6.show()
fig6.write_html("analysis6_basket_distribution.html")

# TODO: Additional basket analyses
# - Large vs small basket patterns
# - Basket size trends over time



print("✓ Analysis 6 complete")

# ==============================================================================
# ANALYSIS 7: PRICING ANALYSIS
# ==============================================================================

print("\n" + "="*70)
print(" "*20 + "ANALYSIS 7: PRICING ANALYSIS")
print("="*70)

# Price distribution
fig7a = px.histogram(
    df_clean[df_clean['UnitPrice'] < df_clean['UnitPrice'].quantile(0.95)],
    x='UnitPrice',
    nbins=50,
    title='Product Price Distribution (excluding top 5%)',
    labels={'UnitPrice': 'Unit Price ($)'}
)
fig7a.update_layout(height=500)
fig7a.show()
fig7a.write_html("analysis7_price_distribution.html")

# Revenue by price range
df_clean['PriceRange'] = pd.cut(df_clean['UnitPrice'], 
                                bins=[0, 5, 10, 20, 50, 100, float('inf')],
                                labels=['$0-5', '$5-10', '$10-20', '$20-50', '$50-100', '$100+'])

price_range_revenue = df_clean.groupby('PriceRange')['TotalPrice'].sum()

fig7b = px.bar(
    x=price_range_revenue.index.astype(str),
    y=price_range_revenue.values,
    title='Revenue by Price Range',
    labels={'x': 'Price Range', 'y': 'Revenue ($)'},
    color=price_range_revenue.values,
    color_continuous_scale='Greens'
)
fig7b.update_layout(height=500, showlegend=False)
fig7b.show()
fig7b.write_html("analysis7_price_ranges.html")

# TODO: Additional pricing analyses
# - Price elasticity insights
# - Optimal price points



print("✓ Analysis 7 complete")

# ==============================================================================
# ANALYSIS 8: RETURNS/CANCELLATIONS
# ==============================================================================

print("\n" + "="*70)
print(" "*20 + "ANALYSIS 8: RETURNS & CANCELLATIONS")
print("="*70)

# Identify returns (negative quantities)
returns = df_clean[df_clean['Quantity'] < 0]
sales = df_clean[df_clean['Quantity'] > 0]

print(f"\n↩️ RETURN STATISTICS:")
print(f"Total return transactions: {returns['InvoiceNo'].nunique():,}")
print(f"Total items returned: {abs(returns['Quantity'].sum()):,.0f}")
print(f"Return rate: {len(returns)/len(df_clean)*100:.2f}%")
print(f"Financial impact: ${abs(returns['TotalPrice'].sum()):,.2f}")

# Top returned products
if len(returns) > 0:
    top_returned = returns.groupby('Description')['Quantity'].sum().nsmallest(10)
    
    fig8 = px.bar(
        x=abs(top_returned.values),
        y=top_returned.index,
        orientation='h',
        title='Top 10 Most Returned Products',
        labels={'x': 'Quantity Returned', 'y': 'Product'},
        color=abs(top_returned.values),
        color_continuous_scale='Reds'
    )
    fig8.update_layout(height=500, showlegend=False)
    fig8.show()
    fig8.write_html("analysis8_top_returns.html")

# TODO: Additional return analyses
# - Return patterns over time
# - Countries with highest return rates



print("✓ Analysis 8 complete")

# ==============================================================================
# ANALYSIS 9: COHORT ANALYSIS
# ==============================================================================

print("\n" + "="*70)
print(" "*20 + "ANALYSIS 9: COHORT ANALYSIS")
print("="*70)

# TODO: Complete cohort analysis
# Steps:
# 1. Identify first purchase date for each customer
# 2. Group customers by cohort (month of first purchase)
# 3. Calculate retention rates
# 4. Visualize cohort behavior over time

# Starter code:
# first_purchase = df_clean.groupby('CustomerID')['InvoiceDate'].min()
# ...

print("⚠️ TODO: Complete cohort analysis")

# ==============================================================================
# ANALYSIS 10: FORECASTING & PREDICTIONS
# ==============================================================================

print("\n" + "="*70)
print(" "*20 + "ANALYSIS 10: FORECASTING & PREDICTIONS")
print("="*70)

# Simple trend analysis and projection
# TODO: Create revenue forecast for next quarter
# Can use simple trend line or more advanced methods

# Example: Calculate growth rate
revenue_by_month = df_clean.groupby(df_clean['InvoiceDate'].dt.to_period('M'))['TotalPrice'].sum()
growth_rates = revenue_by_month.pct_change()

print(f"\n📈 AVERAGE MONTHLY GROWTH RATE: {growth_rates.mean()*100:.2f}%")

# TODO: Create forecast visualization
# Project next 3 months based on trend



print("✓ Analysis 10 complete")

# ==============================================================================
# SUMMARY STATISTICS
# ==============================================================================

print("\n" + "="*70)
print(" "*25 + "PROJECT SUMMARY")
print("="*70)

summary_stats = {
    'Total Revenue': f"${total_revenue:,.2f}",
    'Total Transactions': f"{total_transactions:,}",
    'Total Customers': f"{total_customers:,}",
    'Average Order Value': f"${avg_order_value:,.2f}",
    'Top Country': country_revenue.index[0],
    'Date Range': f"{df_clean['InvoiceDate'].min()} to {df_clean['InvoiceDate'].max()}",
    'Total Products': f"{df_clean['StockCode'].nunique():,}",
    'Countries Served': f"{df_clean['Country'].nunique():,}"
}

summary_df = pd.DataFrame(list(summary_stats.items()), columns=['Metric', 'Value'])
summary_df.to_csv("project_summary.csv", index=False)

print("\n✓ Summary statistics saved to project_summary.csv")

# ==============================================================================
# NEXT STEPS
# ==============================================================================

print("\n" + "="*70)
print(" "*28 + "NEXT STEPS")
print("="*70)

print("""
📋 REMAINING TASKS:

1. COMPLETE ALL ANALYSES
   [ ] Analysis 1-8 are started - refine and enhance
   [ ] Complete Analysis 9 (Cohort Analysis)
   [ ] Complete Analysis 10 (Forecasting)

2. CREATE COMPREHENSIVE DASHBOARD
   [ ] Decide on dashboard format (Plotly Dash, Streamlit, HTML, etc.)
   [ ] Integrate all visualizations
   [ ] Add executive summary
   [ ] Polish design and layout

3. CREATE POWERPOINT PRESENTATION
   [ ] Title slide with team info
   [ ] Business problem (1-2 slides)
   [ ] Methodology (1 slide)
   [ ] Key visualizations (6-8 slides)
   [ ] Insights & recommendations (2-3 slides)
   [ ] Limitations & future work (1 slide)

4. WRITE TECHNICAL DOCUMENTATION
   [ ] Executive summary
   [ ] Data preparation details
   [ ] Analytical approach
   [ ] Key findings
   [ ] Business recommendations
   [ ] Limitations

5. FINALIZE CODE
   [ ] Add comments throughout
   [ ] Create README.md
   [ ] Create requirements.txt
   [ ] Test that everything runs
   [ ] Organize files logically

6. QUALITY CHECK
   [ ] All visualizations display correctly
   [ ] All files are properly named
   [ ] ZIP file is ready for submission
   [ ] Team review complete

Good luck! 🚀
""")

print(f"\nEnd Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print("="*70)
