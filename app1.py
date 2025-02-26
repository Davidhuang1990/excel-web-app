import pandas as pd
import numpy as np
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import OneHotEncoder

# Validation mappings (using the subset from your latest code)
VALIDATION_MAPPINGS = {
    "Plastic": {
        "Packaging Form": ["Product packaging (flexible)", "Product packaging (rigid, excluding beverage bottle)", "Transport and protective packaging"],
        "Further Details": ["HDPE", "LDPE", "Others"]
    },
    "Paper": {
        "Packaging Form": ["Product packaging (rigid, excluding beverage bottle)", "Transport and protective packaging"],
        "Further Details": ["Paper", "Corrugated board"]
    },
    "Metal": {
        "Packaging Form": ["Product packaging (rigid, excluding beverage bottle)"],
        "Further Details": ["Steel"]
    },
    "Wood": {
        "Packaging Form": ["Transport and protective packaging"],
        "Further Details": ["N/A"]
    },
    "Glass": {
        "Packaging Form": ["Product packaging (excluding beverage bottle)"],
        "Further Details": ["Green", "Clear"]
    },
    "Composite": {
        "Packaging Form": ["Beverage carton"],
        "Further Details": []
    },
    "Others": {
        "Packaging Form": ["Others"],
        "Further Details": ["Biodegradable/compostable"]
    }
}

# Generate 200 rows of synthetic data
np.random.seed(42)  # For reproducibility
n_samples = 200
materials = list(VALIDATION_MAPPINGS.keys())
data = {
    "Packaging Material": np.random.choice(materials, n_samples, p=[0.4, 0.2, 0.15, 0.15, 0.05, 0.03, 0.02]),  # Bias towards Plastic
    "Packaging Form": [],
    "Further Details": [],
    "Weight (kg)": []
}

for i in range(n_samples):
    material = data["Packaging Material"][i]
    forms = VALIDATION_MAPPINGS[material]["Packaging Form"]
    data["Packaging Form"].append(np.random.choice(forms))
    
    details = VALIDATION_MAPPINGS[material]["Further Details"]
    if details:
        data["Further Details"].append(np.random.choice(details))
    else:
        data["Further Details"].append("N/A" if material == "Wood" else np.nan)
    
    # Weights: Log-normal distribution (2 to 200,000 kg)
    weight = np.random.lognormal(mean=7, sigma=2)
    data["Weight (kg)"].append(max(2.0, min(weight, 200000.0)))

synthetic_df = pd.DataFrame(data)

# Save to CSV
csv_file = "synthetic_historical_data.csv"
synthetic_df.to_csv(csv_file, index=False)
print(f"Synthetic data saved to {csv_file}")

# Train the model and get predictions
def train_and_predict(df):
    features = ["Packaging Material", "Packaging Form", "Further Details"]
    target = "Weight (kg)"
    
    df = df.dropna(subset=[target] + features)
    if df.empty:
        return df, []
    
    X = df[features]
    y = df[target]
    
    encoder = OneHotEncoder(sparse_output=False, handle_unknown='ignore')
    X_encoded = encoder.fit_transform(X)
    
    model = RandomForestRegressor(n_estimators=100, random_state=42)
    model.fit(X_encoded, y)
    
    y_pred = model.predict(X_encoded)
    df["Predicted Weight (kg)"] = y_pred
    return df, list(y_pred)

# Get predictions
synthetic_df_with_preds, predicted_weights = train_and_predict(synthetic_df)

# Group by unique combinations and calculate mean predicted weight
unique_combinations = synthetic_df_with_preds.groupby(
    ["Packaging Material", "Packaging Form", "Further Details"]
)["Predicted Weight (kg)"].mean().reset_index()

# Display synthetic data sample
print("\nSample of Synthetic Data (first 10 rows):")
print(synthetic_df_with_preds.head(10))

# Display predicted weights for unique combinations
print("\nPredicted Weights for Each Unique Combination:")
for idx, row in unique_combinations.iterrows():
    print(f"Combination {idx + 1}: {row['Packaging Material']}, {row['Packaging Form']}, {row['Further Details']} "
          f"-> Predicted Weight = {row['Predicted Weight (kg)']:.2f} kg")
