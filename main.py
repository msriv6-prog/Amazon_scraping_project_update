import pandas as pd
import requests
from bs4 import BeautifulSoup
import lxml
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# scraping function
def scrape_amazon():
    urls = url_entry.get("1.0", "end").strip().splitlines() # multiple urls from text widget
    if not urls:
        messagebox.showerror("Error", "Please enter one or more Amazon urls")
        return
    headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152.0.0.0 Safari/537.36 Edg/152.0.0.0"
     }
    #  containers for batch results
    all_product_info, all_additional, all_reviews, all_details1, all_details2 = [], [], [], [], []
    for idx, url in enumerate(urls, start=1):
        try:
            response = requests.get(url, headers=headers)
            if response.status_code != 200:
                    messagebox.showerror("Error", f"Fetching Error - {response.status_code} for {url}")
                    continue
                
            soup = BeautifulSoup(response.text, "lxml")

    
            product_title = soup.find("span", id="productTitle").text.strip()
            product_price = soup.find("span", class_="a-price-whole").text.strip()
            span_tag = soup.find("span", id="acrPopover")
            product_rating = span_tag.get("title")
            product_details = soup.find("ul", class_="a-unordered-list a-vertical a-spacing-mini").text.strip()
            product_reviews = soup.find("ul", id="localTopReviewsList").text.strip()
            data = {}
            table = soup.find("table", class_="a-normal a-spacing-micro")
            if table:
                rows = table.find_all("tr")
                for row in rows:
                    cells = row.find_all("td")
                    if len(cells) == 2:
                        key = cells[0].text.strip()
                        value = cells[1].text.strip()
                        data[key] = value
            review_content = soup.find("ul", id="localTopReviewsList")
            reviews = []
            if review_content:
                customer_names = [name.text.strip() for name in review_content.find_all("span", class_="a-profile-name")]
                customer_ratings = [rating.text.strip() for rating in review_content.find_all("span", class_="a-icon-alt")]
                customer_comments = [comment.text.strip() for comment in review_content.find_all("h5", class_="_Y3Itd_single-review-title_2aKRE")]
                customer_additionals = [additional.text.strip() for additional in review_content.find_all("div", class_="_Y3Itd_contain-rich-content_2IORW")]
                # Combine all into rows
        
                for name, rating, comment, additional in zip(customer_names, customer_ratings, customer_comments, customer_additionals):
                        reviews.append({
                        "Product URL" : url,    
                        "Customer Name": name,
                        "Rating": rating,
                        "Comment": comment,
                        "Additional": additional
                        })
            table_list = soup.find_all("table", class_="a-keyvalue prodDetTable")
            data1, data2 = {}, {}
            for idx, table1 in enumerate(table_list[:5]):
                 for row1 in table1.find_all("tr"):
                      cells1 = row1.find_all(["th", "td"])
                      if len(cells1) == 2:
                           data1[cells1[0].text.strip()] = cells1[1].text.strip()
            for table2 in table_list[6:10]:
                 for row2 in table2.find_all("tr"):
                      cells2 = row2.find_all(["th", "td"])
                      if len(cells2) == 2:
                           data2[cells2[0].text.strip()] = cells2[1].text.strip()               
                 
            product_info1 = [{
            "Product_title": product_title,
            "Product_price": product_price,
            "Product_rating": product_rating,
            "Product_details": product_details,
            "Product_reviews": product_reviews
            }]
            # Convert to DataFrame for column-wise output
            df12 = pd.DataFrame(product_info1)
            df9 = pd.DataFrame(list(data.items()), columns=["Attribute", "Value"])
            df = pd.DataFrame(reviews)
            df4 = pd.DataFrame(list(data1.items()), columns=["Attribute", "Value"])
            df5 = pd.DataFrame(list(data2.items()), columns=["Attribute", "Value"])
            progress_bar['value'] = (idx/len(urls))*100
            root.update_idletasks()
        except Exception as e:
            messagebox.showerror("Error", f"Parsing error for {url}: {e}")
            continue   
    

    # save to excel
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Save output file as")
    if save_path:
         with pd.ExcelWriter(save_path) as writer:
                 df12.to_excel(writer, sheet_name="Product_Info", index=False)
                 df9.to_excel(writer, sheet_name="Additional_Details", index=False)
                 df.to_excel(writer, sheet_name="Product_Reviews", index=False)
                 df4.to_excel(writer, sheet_name="Laptop_Details1", index=False)
                 df5.to_excel(writer, sheet_name="Laptop_Details2", index=False)
    else:
         messagebox.showinfo("Cancelled", "Save operation cancelled")

              
# Update GUI
    result_label.config(text=f"✅ Scraping complete!\nSaved as {save_path}")
    progress_bar['value'] = 100    
             
    
    
# Tkinter GUI setup
root = tk.Tk()
root.title("Amazon Product Scraper")
root.geometry("600x400")

# url input
tk.Label(root, text="Enter Amazon Product urls (one per line):", font=("Arial", 12)).pack(pady=10)
url_entry = tk.Text(root, width=70, height=8) # multi line input
url_entry.pack(pady=5)

# Scrape button
scrape_button = tk.Button(root, text="Scrape Data", command=scrape_amazon, bg="green", fg="white", font=("Arial", 12))
scrape_button.pack(pady=10)

# progress bar
progress_bar = ttk.Progressbar(root, length=500, mode="determinate")
progress_bar.pack(pady=10)


# result label
result_label = tk.Label(root, text="", font=("Arial", 12), fg="blue")
result_label.pack(pady=10)

root.mainloop()




























