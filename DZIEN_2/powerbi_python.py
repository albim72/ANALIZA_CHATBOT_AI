import matplotlib.pyplot as plt

grupa = dataset.groupby("kategoria")["wartosc"].sum().sort_values()

plt.figure(figsize=(10, 6))
grupa.plot(kind="barh", color="#2C7BE5")
plt.title("Sprzedaż wg kategorii (Python + Power BI)")
plt.xlabel("Wartość sprzedaży")
plt.ylabel("")
plt.tight_layout()
plt.show()
