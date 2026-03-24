import pandas as pd

# Data contoh
data = {
    "ID": ["P001", "P002", "P003"],
    "Nama_Produk": ["Laptop Asus A15", "Mouse Logitech M185", "Keyboard Mechanical"],
    "Kategori": ["Elektronik", "Aksesoris", "Aksesoris"],
    "Harga": [8500000, 150000, 450000],
    "Stok": [12, 85, 30],
    "Rating": [4.5, 4.2, 4.8],
}

df = pd.DataFrame(data)
df.to_excel("data/input/contoh_produk.xlsx", index=False)
print("File contoh dibuat: data/input/contoh_produk.xlsx")
