import pandas as pd
from datetime import datetime, timedelta

# Starting from the last period in the provided example
start_date = datetime.strptime("21/Dec/2020", "%d/%b/%Y")
end_date = datetime.strptime("20/Jan/2021", "%d/%b/%Y")
destin_period = 202101

# List to store new rows
new_rows = []

# Generate rows until December 2026
while destin_period <= 202612:
    origin_period = f"{start_date.strftime('%d/%b/%Y')} - {end_date.strftime('%d/%b/%Y')}"
    new_rows.append({"Origin_period": origin_period, "Destin_period": destin_period})

    # Increment to next month
    start_date = end_date + timedelta(days=1)
    end_date = (start_date + timedelta(days=31)).replace(day=20)  # Ensuring it reaches the 20th of the next month
    destin_period += 1

# Convert to DataFrame
extended_df = pd.DataFrame(new_rows)

# Save DataFrame to CSV
output_file_path = 'C:\dhis2\Extended_Periods_2026.csv'
extended_df.to_csv(output_file_path, index=False)

output_file_path
