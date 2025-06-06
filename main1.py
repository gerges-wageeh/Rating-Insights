import pandas as pd
import matplotlib.pyplot as plt

# Load customer reviews dataset
data = pd.read_csv('customer_reviews.csv')

# Initial data exploration
print('Display five lines:')
print(f'\n{data.head()}\n')  # Display the first 5 rows of the dataset
print('Column information:')
print(f'\n{data.info()}\n')  # Print column data types and non-null counts

# Total rating (sum of stars) received by each product
# This helps identify which products received the most attention overall
product_evaluation = (
    data.groupby('product_name')['rating']
    .sum()
    .sort_values(ascending=False)
    .reset_index()
)

# Average rating for each product
# Useful to assess customer satisfaction regardless of the number of reviews
average_product_rating = (
    data.groupby('product_name')['rating']
    .mean()
    .sort_values(ascending=False)
    .round(2)
    .reset_index()
)

# Prepare date column for time-based analysis
data['date'] = pd.to_datetime(data['date'])         # Convert string to datetime
data['month'] = data['date'].dt.to_period('M')      # Extract month for aggregation

# Monthly trend of ratings (are ratings increasing or dropping over time?)
monthly_ratings = data.groupby('month')['rating'].sum()
monthly_ratings.index = monthly_ratings.index.astype(str)  # Convert index for plotting

# Plot ratings over time
ax = monthly_ratings.plot(marker="o", figsize=(10, 5))
plt.title("Evaluations over time")
plt.xlabel("Month")
plt.ylabel("Total Ratings")
plt.grid(True)
plt.tight_layout()
plt.xticks(rotation=45)
plt.savefig("reviews.png", bbox_inches="tight", dpi=300)  # Save figure as high-res image
plt.show()

# Filter products with rating > 3 (high-rated products)
most_rated_products = data[(data['rating'] > 3)].reset_index(drop=True)

# Prepare Excel report with three sheets: filtered products, averages, total ratings

# Convert date to string format for easier reading in Excel
most_rated_products['date'] = most_rated_products['date'].dt.strftime('%Y-%m-%d')

# Write results to Excel file with formatted column widths
with pd.ExcelWriter("Analyze product reviews.xlsx", engine='xlsxwriter') as writer:
    sheets_and_data = {
        "Most_rated_products": most_rated_products,
        "average_product_rating": average_product_rating,
        "product_evaluation": product_evaluation
    }

    for sheet_name, df in sheets_and_data.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)  # Export each dataframe to its sheet

        worksheet = writer.sheets[sheet_name]  # Access the corresponding Excel worksheet

        # Adjust column width based on content length for better readability
        for i, col in enumerate(df.columns):
            max_len = max(
                df[col].astype(str).map(len).max(),  # Longest cell value in the column
                len(str(col))                        # Length of the column name
            ) + 2
            worksheet.set_column(i, i, max_len)     # Set width for column i
