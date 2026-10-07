"""Kurgusal demo proje / Fictional demo project: pandas -> Excel reports."""
import argparse
from decimal import Decimal, InvalidOperation
from pathlib import Path
import pandas as pd

REQUIRED = ["order_id", "date", "product", "quantity", "unit_price"]
COPY = {
    "tr":{"summary":"Özet", "sales":"Satışlar", "title":"Ürün bazında satış raporu", "demo":"Kurgusal demo proje. Tutarlar TL; vergi ve indirim hariç.", "total":"Brüt satış (TL)", "units":"Toplam adet", "headers":["Ürün", "Adet", "Brüt satış (TL)"], "raw":["Kayıt kimliği", "Tarih", "Ürün", "Adet", "Birim fiyat (TL)", "Brüt satış (TL)"], "chart":"Ürün bazında brüt satış (TL)"},
    "en":{"summary":"Summary", "sales":"Sales", "title":"Sales report by product", "demo":"Fictional demo project. TRY amounts; before tax and discounts.", "total":"Gross sales (TRY)", "units":"Total units", "headers":["Product", "Units", "Gross sales (TRY)"], "raw":["Entry ID", "Date", "Product", "Units", "Unit price (TRY)", "Gross sales (TRY)"], "chart":"Gross sales by product (TRY)"},
}


def load_sales(source):
    frame = pd.read_csv(source, dtype=str, keep_default_na=False)
    if frame.empty or not set(REQUIRED).issubset(frame.columns):
        raise ValueError("Non-empty CSV with required columns needed / Gerekli kolonları olan dolu CSV gerekir")
    frame = frame[REQUIRED].copy()
    for column in ("order_id", "product"):
        frame[column] = frame[column].str.strip()
        if frame[column].eq("").any():
            raise ValueError(f"Blank {column} / Boş alan")
    if frame["order_id"].duplicated().any():
        raise ValueError("Duplicate entry ID / Tekrarlanan kayıt kimliği")
    frame["date"] = pd.to_datetime(frame["date"], format="%Y-%m-%d", errors="raise")
    quantities = pd.to_numeric(frame["quantity"], errors="raise")
    if ((quantities <= 0) | (quantities > 1_000_000) | (quantities % 1 != 0)).any():
        raise ValueError("Quantity must be a positive bounded integer / Adet pozitif tam sayı olmalı")
    frame["quantity"] = quantities.astype("int64")
    def cents(value):
        try:
            amount = Decimal(value)
            if not amount.is_finite() or amount < 0 or amount > 1_000_000 or amount.as_tuple().exponent < -2:
                raise ValueError("Invalid price / Geçersiz fiyat")
            return int(amount * 100)
        except InvalidOperation:
            raise ValueError("Invalid price / Geçersiz fiyat") from None
    frame["price_cents"] = frame["unit_price"].map(cents)
    frame["gross_cents"] = frame["quantity"] * frame["price_cents"]
    return frame


def analyze(frame):
    return frame.groupby("product", as_index=False).agg(quantity=("quantity","sum"), gross_cents=("gross_cents","sum")).sort_values("gross_cents", ascending=False, kind="stable").reset_index(drop=True)


def write_report(frame, destination, language):
    text, grouped = COPY[language], analyze(frame)
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(destination, engine="xlsxwriter", engine_kwargs={"options":{"strings_to_formulas":False, "strings_to_urls":False}}) as writer:
        workbook = writer.book
        summary = workbook.add_worksheet(text["summary"])
        sales = workbook.add_worksheet(text["sales"])
        basic = workbook.add_format({"font_name":"Arial", "font_size":10, "valign":"vcenter"})
        title = workbook.add_format({"font_name":"Arial", "font_size":14, "bold":True, "font_color":"#243E32"})
        note = workbook.add_format({"font_name":"Arial", "font_size":10, "italic":True, "font_color":"#637066"})
        head = workbook.add_format({"font_name":"Arial", "font_size":10, "bold":True,"bg_color":"#243E32","font_color":"white","align":"center","valign":"vcenter"})
        money = workbook.add_format({"font_name":"Arial", "font_size":10, "num_format":'#,##0.00', "valign":"vcenter"})
        integer = workbook.add_format({"font_name":"Arial", "font_size":10, "num_format":'#,##0', "valign":"vcenter"})
        date = workbook.add_format({"font_name":"Arial", "font_size":10, "num_format":'yyyy-mm-dd', "valign":"vcenter"})
        for sheet in (summary,sales):
            sheet.hide_gridlines(2)
            sheet.set_default_row(22)
            sheet.set_landscape()
            sheet.fit_to_pages(1,0)
        summary.set_tab_color("#243E32")
        summary.set_column("A:A",34,basic); summary.set_column("B:B",14,integer); summary.set_column("C:C",22,money); summary.set_column("D:D",3); summary.set_column("E:L",11,basic)
        sales.set_column("A:A",19,basic); sales.set_column("B:B",15,date); sales.set_column("C:C",34,basic); sales.set_column("D:D",12,integer); sales.set_column("E:F",22,money)
        summary.write("A2",text["title"],title); summary.write("A3",text["demo"],note)
        sales.write("A2",text["sales"],title)
        sales.write("A3","Source / Kaynak: data/demo_sales.csv" if Path(frame.attrs.get("source","")).name == "demo_sales.csv" else "Source / Kaynak: input CSV",note)
        sales.write_row(4,0,text["raw"],head)
        for index,row in enumerate(frame.itertuples(index=False),5):
            excel_row=index+1
            sales.write_string(index,0,row.order_id,basic)
            sales.write_datetime(index,1,row.date.to_pydatetime(),date)
            sales.write_string(index,2,row.product,basic)
            sales.write_number(index,3,int(row.quantity),integer)
            sales.write_number(index,4,int(row.price_cents)/100,money)
            sales.write_formula(index,5,f"=D{excel_row}*E{excel_row}",money,int(row.gross_cents)/100)
        end=len(frame)+5
        quoted="'"+text["sales"]+"'"
        summary.write("A4",text["total"],basic); summary.write_formula("B4",f"=SUM({quoted}!F6:F{end})",money,int(frame["gross_cents"].sum())/100)
        summary.write("A5",text["units"],basic); summary.write_formula("B5",f"=SUM({quoted}!D6:D{end})",integer,int(frame["quantity"].sum()))
        summary.write_row(6,0,text["headers"],head)
        for index,row in enumerate(grouped.itertuples(index=False),7):
            excel_row=index+1
            summary.write_string(index,0,row.product,basic)
            summary.write_formula(index,1,f"=SUMIFS({quoted}!D6:D{end},{quoted}!C6:C{end},A{excel_row})",integer,int(row.quantity))
            summary.write_formula(index,2,f"=SUMIFS({quoted}!F6:F{end},{quoted}!C6:C{end},A{excel_row})",money,int(row.gross_cents)/100)
        chart=workbook.add_chart({"type":"bar"})
        chart.add_series({"name":text["total"],"categories":[text["summary"],7,0,6+len(grouped),0],"values":[text["summary"],7,2,6+len(grouped),2],"fill":{"color":"#477667"},"border":{"none":True}})
        chart.set_title({"name":text["chart"],"name_font":{"name":"Arial","size":11}})
        chart.set_legend({"none":True}); chart.set_x_axis({"min":0,"num_format":"#,##0","num_font":{"name":"Arial","size":10}}); chart.set_y_axis({"reverse":True,"num_font":{"name":"Arial","size":10}})
        summary.insert_chart("E7",chart,{"x_scale":1.2,"y_scale":1.05})
        sales.freeze_panes(5,0);sales.autofilter(4,0,end-1,5)
        summary.print_area(0,0,max(22,8+len(grouped)),11)
    return grouped


def main():
    parser=argparse.ArgumentParser(description="Sales Excel automation / Satış Excel otomasyonu")
    parser.add_argument("--input",type=Path,default=Path("data/demo_sales.csv"))
    parser.add_argument("--output",type=Path,default=Path("reports"))
    args=parser.parse_args()
    try:
        frame=load_sales(args.input);frame.attrs["source"]=str(args.input)
        for language in COPY:
            write_report(frame,args.output/f"sales_{language}.xlsx",language)
        print(f"Rows / Satırlar: {len(frame)}; gross TRY / brüt TL: {int(frame['gross_cents'].sum())/100:.2f}")
    except (ValueError,OSError) as exc:
        parser.exit(1,f"Failed / Başarısız: {exc}\n")


if __name__=="__main__":
    main()
