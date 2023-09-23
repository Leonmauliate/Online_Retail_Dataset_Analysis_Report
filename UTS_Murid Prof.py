## Import Library yang Akan Digunakan

import pandas as pd
import numpy


## Pembuatan Function-function

# Function untuk Extract data dari File Excel
def extract_excel(file_path, sheet):
    data = pd.read_excel(file_path, sheet_name = sheet)
    return data

# Function untuk Mengubah Nama Kolom
def change_col_name(data):
    data = data.rename(columns={"Invoice" : "invoice", "StockCode" : "stock_code", "Description" : "description",
                                "Quantity" : "quantity", "InvoiceDate" : "invoice_date", "Price" : "price",
                                "Customer ID" : "customer_id", "Country" : "country"})
    return data

# Function untuk Menggabungkan Dua Buah Data (data pertama lalu dilanjuti data kedua)
def merge_data(data_a, data_b):
    frames = [data_a, data_b]
    merge = pd.concat(frames, ignore_index = True)
    return merge

# Function untuk Mengubah Tipe Data dari Suatu Kolom
def convert_data_type(data, column, data_type):
    data[column] = data[column].astype(data_type)
    return data

# Function untuk Menghapus Baris Data Duplikat
def remove_duplicate(data):
    # Menghilangkan semua baris data yang sama kecuali baris data pertama
    data = data.drop_duplicates(keep='first', ignore_index = True) 
    return data

# Function untuk Menghapus Baris Data yang Mengandung Informasi yang Tidak Lengkap
def remove_missing_value(data):
    data = data.dropna().reset_index(drop = True)
    return data

# Function untuk Menghapus Baris Data yang Memiliki "quantity" Negatif
def remove_negative_qty(data):
    temp = []
    for row in range(len(data.index)):
        if data["quantity"][row]<0:
            temp.append(row)
    data_clean = data.drop(temp).reset_index(drop = True)
    return data_clean

# Function untuk Menghitung total_order, yaitu Hasil Perkalian "quantity" dengan "price"
def total_order(data):
    total = []
    for row in range(len(data.index)):
        temp = data["quantity"][row] * data["price"][row]
        total.append(temp)
    data["total_order"] = total
    return data

# Function untuk Mencari Produk Terlaris
def produk_terlaris(data, ukuran):
    # Ukuran dapat berupa "quantity" atau "total_order"
    code_group = data[["stock_code", ukuran]].groupby("stock_code").sum()
    merge = pd.merge(code_group, df_produk, on="stock_code")
    return merge.sort_values(ukuran, ascending=False).head()

# Function untuk Mengurutkan Customer
def pelanggan_terbanyak(data, ukuran):
    # Ukuran dapat berupa "quantity" atau "total_order"
    cust_group = data[["customer_id", ukuran]].groupby("customer_id").sum()
    merge = pd.merge(cust_group, df_customer, on="customer_id")
    return merge.sort_values(ukuran, ascending=False)

# Function untuk Melakukan Klasifikasi atau Segmentasi Pelanggan
def segmentasi_pelanggan(data, quartile, ukuran):
    # Ukuran dapat berupa "quantity" atau "total_order"
    segmentasi = []
    for row in range(len(data.index)):
        if data[ukuran][row] <= quartile[0]:
            segmentasi.append("small")
        elif data[ukuran][row] <= quartile[1]:
            segmentasi.append("medium")
        elif data[ukuran][row] <= quartile[2]:
            segmentasi.append("high")
        else:
            segmentasi.append("very high")
    data["cluster"] = segmentasi
    return data

# Function untuk Mencari Produk Terlaris pada suatu Triwulan atau Kuartal
def trend_quarter(kuartal):
    c = df_penjualan[df_penjualan["quarter"] == kuartal]
    c1 = produk_terlaris(c, "quantity")
    c2 = produk_terlaris(c, "total_order")
    return c, c1, c2

# Function untuk Menghitung Jumlah Order pada Setiap Jam
def order_per_jam(data):
    order_jam = data["hours"].value_counts()
    return order_jam

# Function untuk Menghitung Jumlah Transaksi Dalam dan Luar Negeri
def dalam_luar(data):
    luar_negeri = 0
    dalam_negeri = 0
    for i in range(len(data)):
        if data[i] == "United Kingdom":
            dalam_negeri = dalam_negeri + 1
        else:
            luar_negeri = luar_negeri + 1
    return luar_negeri, dalam_negeri


## Proses Extract
path = r"C:/Users/ASUS/OneDrive/Documents/Keperluan UNPAR/Semester 6/Kapita Selekta Statistika/UTS/online_retail_II.xlsx"

data_1 = extract_excel(path, "Year 2009-2010")
data_2 = extract_excel(path, "Year 2010-2011")


## Proses Transform
data_1 = change_col_name(data_1)
data_2 = change_col_name(data_2)

joint_data = merge_data(data_1, data_2)

joint_data = convert_data_type(joint_data, ["invoice", "stock_code", "description", "customer_id", "country"], "string")

joint_data = remove_duplicate(joint_data)

data_nonan = remove_missing_value(joint_data)

data_noneg = remove_negative_qty(data_nonan)

data_noneg = convert_data_type(data_noneg, "quantity", "int")

data_full = total_order(data_noneg)

df_customer = remove_duplicate(data_full[["customer_id", "country"]])
df_produk = remove_duplicate(data_full[["stock_code", "description"]])
df_penjualan = data_full[["invoice", "invoice_date", "customer_id", "stock_code", "quantity", "price", "total_order"]]
df_penjualan_awal = df_penjualan


# Produk Terlaris berdasarkan quantity
prod_laris_qty = produk_terlaris(df_penjualan, "quantity")
print(prod_laris_qty)

# Produk Terlaris berdasarkan total_order
prod_laris_total = produk_terlaris(df_penjualan, "total_order")
print(prod_laris_total)

# Segmentasi Pelanggan berdasarkan total_order
segmen_1 = pelanggan_terbanyak(df_penjualan, "total_order").reset_index(drop = True)
quartile_1 = numpy.quantile(segmen_1["total_order"], [0.25,0.5,0.75])
segmen_total_order = segmentasi_pelanggan(segmen_1, quartile_1, "total_order")

# Segmentasi Pelanggan berdasarkan quantity
segmen_2 = pelanggan_terbanyak(df_penjualan, "quantity").reset_index(drop = True)
quartile_2 = numpy.quantile(segmen_2["quantity"], [0.25,0.5,0.75])
segmen_quantity = segmentasi_pelanggan(segmen_2, quartile_2, "quantity")


# Produk terlaris (per Kuartal Tahun)
df_penjualan["quarter"] = df_penjualan["invoice_date"].dt.quarter

# Kuartal 1
quarter1 = trend_quarter(1)[1:3]  # Produk terlaris pada kuartal 1 (quantity dan total_order)
print(quarter1)
c1 = trend_quarter(1)[0]
c1_final = pd.merge(c1, df_produk, on="stock_code")

# Kuartal 2
quarter2 = trend_quarter(2)[1:3]  # Produk terlaris pada kuartal 2 (quantity dan total_order)
print(quarter2)
c2 = trend_quarter(2)[0]
c2_final = pd.merge(c2, df_produk, on="stock_code")

# Kuartal 3
quarter3 = trend_quarter(3)[1:3]  # Produk terlaris pada kuartal 3 (quantity dan total_order)
print(quarter3)
c3 = trend_quarter(3)[0]
c3_final = pd.merge(c3, df_produk, on="stock_code")

# Kuartal 4
quarter4 = trend_quarter(4)[1:3]  # Produk terlaris pada kuartal 4 (quantity dan total_order)
print(quarter4)
c4 = trend_quarter(4)[0]
c4_final = pd.merge(c4, df_produk, on="stock_code")


# Jumlah Order per Jam
df_penjualan["hours"] = df_penjualan["invoice_date"].dt.hour

banyak_order = order_per_jam(df_penjualan)
print(banyak_order)
df_banyak_order = pd.DataFrame({
    'time_in_hours': ['12', '13', '14', '11', '15',
                      '10', '16', '9', '17', '8',
                      '19', '18', '20', '7', '6'],
    'orders_count': [138995, 126957, 106717, 96858, 86151,
                     72308, 51109, 40984, 26246, 15183,
                     7897, 7179, 1857, 1053, 1]
})


# Penjualan Dalam dan Luar Negeri
cust_penjualan = pd.merge(df_penjualan, df_customer, on = "customer_id")
country = cust_penjualan["country"]

banyak_transaksi = dalam_luar(country)
print(banyak_transaksi)
df_banyak_transaksi = pd.DataFrame({
    'transaksi': ['luar_negeri', 'dalam_negeri'],
    'banyak': [banyak_transaksi[0], banyak_transaksi[1]]
})


## Proses Load
data_full.to_excel("df_full_data_awal.xlsx")

df_customer.to_excel("df_customer.xlsx")
df_produk.to_excel("df_produk.xlsx")
df_penjualan_awal.to_excel("df_penjualan_awal.xlsx")

segmen_total_order.to_excel("segmentasi_total_order.xlsx")
segmen_quantity.to_excel("segmentasi_quantity.xlsx")

c1_final.to_excel("quarter_1.xlsx")
c2_final.to_excel("quarter_2.xlsx")
c3_final.to_excel("quarter_3.xlsx")
c4_final.to_excel("quarter_4.xlsx")

df_banyak_order.to_excel("banyak_order_tiap_jam.xlsx")

df_banyak_transaksi.to_excel("banyak_transaksi_dalam_luar_negeri.xlsx")

penjualan_produk = pd.merge(df_penjualan, df_produk, on="stock_code")
data_full_final = pd.merge(penjualan_produk, df_customer, on="customer_id")

data_full_final.to_excel("df_full_data_final.xlsx")

