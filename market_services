import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load data from the Excel sheets
market_a_calls = pd.read_excel('customerserviceanalyst2.xlsx', sheet_name='MarketA_calls')
market_e_calls = pd.read_excel('customerserviceanalyst2.xlsx', sheet_name='MarketE_calls')
market_a_emails = pd.read_excel('customerserviceanalyst2.xlsx', sheet_name='MarketA_emails')
market_e_emails = pd.read_excel('customerserviceanalyst2.xlsx', sheet_name='MarketE_emails')

# Add 'Market' and 'Source' columns
for df, market, source in [(market_a_calls, 'Market A', 'Calls'), (market_e_calls, 'Market E', 'Calls'), 
                           (market_a_emails, 'Market A', 'Emails'), (market_e_emails, 'Market E', 'Emails')]:
    df['Market'] = market
    df['Source'] = source

# Concatenate all data into a single DataFrame
data = pd.concat([market_a_calls, market_e_calls, market_a_emails, market_e_emails], ignore_index=True)

# Convert Wrap-Up Labels to numeric values if necessary
for col in ['Wrap-Up Label 0', 'Wrap-Up Label 1']:
    data[col] = pd.to_numeric(data[col], errors='coerce')

# Calculate Call and Email Metrics
# For calls: Use Time Ringing, Time Hunting, and estimated wait times
data['Estimated Wait Time (Calls)'] = data['Time Ringing'] + data['From Call Queue Last Wait Seconds'].fillna(0)

# For emails: Calculate response time using Date/Time Opened and Date/Time Closed
data['Date/Time Opened'] = pd.to_datetime(data['Date/Time Opened'], errors='coerce')
data['Date/Time Closed'] = pd.to_datetime(data['Date/Time Closed'], errors='coerce')
data['Email Response Time'] = (data['Date/Time Closed'] - data['Date/Time Opened']).dt.total_seconds()

# Dashboard Title
st.title("Incoming Communication Dashboard")

# Sidebar filters for interactivity
st.sidebar.header("Filter Options")
markets = st.sidebar.multiselect("Select Market(s)", options=data['Market'].unique(), default=data['Market'].unique())
sources = st.sidebar.multiselect("Select Communication Source", options=data['Source'].unique(), default=data['Source'].unique())

# Filter the data based on selections
filtered_data = data[(data['Market'].isin(markets)) & (data['Source'].isin(sources))]

# 1. Overview of Communication Channels by Market
st.subheader("Communication Channels Overview by Market")
channels_count = filtered_data.groupby(['Market', 'Source'])['Source'].count().unstack()
fig1, ax1 = plt.subplots()
channels_count.plot(kind='bar', stacked=True, ax=ax1)
ax1.set_ylabel("Number of Communications")
st.pyplot(fig1)

# 2. Call Wait Time Analysis
st.subheader("Average Estimated Wait Time (Calls) by Market")
if 'Estimated Wait Time (Calls)' in filtered_data.columns:
    wait_time_avg = filtered_data[filtered_data['Source'] == 'Calls'].groupby('Market')['Estimated Wait Time (Calls)'].mean()
    fig2, ax2 = plt.subplots()
    wait_time_avg.plot(kind='bar', color='skyblue', ax=ax2)
    ax2.set_ylabel("Avg Wait Time (seconds)")
    st.pyplot(fig2)

# 3. Email Response Time Analysis
st.subheader("Average Email Response Time by Market")
if 'Email Response Time' in filtered_data.columns:
    email_response_time_avg = filtered_data[filtered_data['Source'] == 'Emails'].groupby('Market')['Email Response Time'].mean()
    fig3, ax3 = plt.subplots()
    email_response_time_avg.plot(kind='bar', color='salmon', ax=ax3)
    ax3.set_ylabel("Avg Response Time (seconds)")
    st.pyplot(fig3)

# 4. Call Types Analysis by Market
st.subheader("Call Types Analysis")
if 'From Call Type' in filtered_data.columns:
    fig4, ax4 = plt.subplots()
    sns.countplot(data=filtered_data, x='From Call Type', hue='Market', ax=ax4)
    ax4.set_xlabel("Call Type")
    ax4.set_ylabel("Count")
    st.pyplot(fig4)

# 5. Wrap-Up Labels Count by Market
st.subheader("Wrap-Up Labels Summary by Market")
wrap_up_labels = filtered_data.groupby('Market')[['Wrap-Up Label 0', 'Wrap-Up Label 1']].sum()
fig5, ax5 = plt.subplots()
wrap_up_labels.plot(kind='bar', ax=ax5, stacked=True)
ax5.set_ylabel("Wrap-Up Labels Count")
st.pyplot(fig5)

# 6. Email Analysis by Market
st.subheader("Email Data Analysis")
if 'Emails' in sources:
    email_counts = filtered_data[filtered_data['Source'] == 'Emails']['Market'].value_counts()
    fig6, ax6 = plt.subplots()
    sns.barplot(x=email_counts.index, y=email_counts.values, ax=ax6)
    ax6.set_xlabel("Market")
    ax6.set_ylabel("Number of Emails")
    st.pyplot(fig6)

# Top Communication Issues by Market as Pie Chart
st.subheader("Top Communication Issues by Market")
if 'Subject' in filtered_data.columns:
    # Aggregate issue counts by market and subject
    issue_counts = filtered_data.groupby(['Market', 'Subject']).size().reset_index(name='Count')
    
    # Select the top 5 issues for each market to keep visuals manageable
    top_issues = issue_counts.sort_values(['Market', 'Count'], ascending=[True, False]).groupby('Market').head(5)

    # Plot each market's top issues as a separate pie chart
    for market in top_issues['Market'].unique():
        st.write(f"Market: {market}")
        market_issues = top_issues[top_issues['Market'] == market]
        
        fig, ax = plt.subplots()
        ax.pie(market_issues['Count'], labels=market_issues['Subject'], autopct='%1.1f%%', startangle=90)
        ax.axis('equal')  # Equal aspect ratio for a circular pie chart
        st.pyplot(fig)


# Display Filtered Data Table
st.subheader("Filtered Data Preview")
st.dataframe(filtered_data)

