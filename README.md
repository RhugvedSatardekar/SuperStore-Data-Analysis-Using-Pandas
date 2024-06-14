
1. **Importing Libraries**:
```python
import pandas as pd
import matplotlib.pyplot as plt
%matplotlib inline
import seaborn as sns
```

2. **Loading Dataset into DataFrame**:
```python
df = pd.read_excel('superstore_sales.xlsx')
```

3. **Data Audit**:
```python
# Check column names
df.columns

# Get DataFrame information
df.info()

# Generate descriptive statistics
df.describe()

# Display first and last 5 rows
df.head(5)
df.tail(5)

# Check for missing values
df.isnull().sum()

# Get shape of DataFrame
df.shape

# Capitalize column names
df.columns = df.columns.str.capitalize()
```

4. **Exploratory Data Analysis**:
   - **Overall Sales Trend**:
   ```python
   df['Order_date'].min()
   df['Order_date'].max()
   ```

   - **Getting Month-Year**:
   ```python
   df['Month-Year'] = df['Order_date'].apply(lambda x: x.strftime('%Y-%m'))
   ```

   - **Grouping Month Year**:
   ```python
   df_trend = df.groupby('Month-Year').sum()['Sales'].reset_index()
   ```

   - **Top 10 Products by Sales**:
   ```python
   prod_sales = pd.DataFrame(df.groupby('Product_name').sum()['Sales'])
   prod_sales.sort_values('Sales', ascending=False).head(10)
   ```

   - **Most Selling Products**:
   ```python
   TopSellingProd = pd.DataFrame(df.groupby('Product_name').sum()['Quantity'])
   TopSellingProd.sort_values('Quantity', ascending=False).head(10)
   ```

   - **Most Preferred Ship Mode**:
   ```python
   df.groupby('Ship_mode').count()['Order_id']
   ```

   - **Most Profitable Category and Sub-Category**:
   ```python
   Highest_Cate_Sub_Cate = pd.DataFrame(df.groupby(['Category', 'Sub_category']).sum()['Profit'])
   Highest_Cate_Sub_Cate.sort_values(['Category', 'Profit'], ascending=False)
   ```

