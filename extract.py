import docx
import time
import datetime
import csv

def parse_docx_tables(docx_path):
    doc = docx.Document(docx_path)
    result = []
    count = 1
    for tbl in doc.tables:
        if len(tbl.rows) < 2:
            continue
        print(f"{count}ページ目を読み取り中")
        count += 1
        for row in range(2, len(tbl.rows)):
            if row % 2 == 0:
                continue
            if tbl.cell(row, 0).text.strip() == "":
                continue
            no = tbl.cell(row, 0).text.strip()
            date = None
            for col in range(1, len(tbl.columns)):
                value = tbl.cell(row, col).text.strip()
                if value != "":
                    try:
                        date = datetime.datetime.strptime(value, '%y/%m/%d')
                    except ValueError:
                        print(f"エラーが発生しました。区域番号 {no} の日付 {value} が間違っています。")
                        date = None
                elif col % 2 == 1 and tbl.cell(row, col - 1).text.strip() != "":
                    date = datetime.datetime.strptime('30/1/1', '%y/%m/%d')
            if date is not None:
                result.append([no, date])
    return result

def write_results_to_csv(results, output_path):
    sorted_result = sorted(results, key=lambda x: x[1])
    with open(output_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f, lineterminator="\n")
        for row in sorted_result:
            n = row[0]
            d = row[1].strftime("%y/%m/%d")
            if d == "30/01/01":
                writer.writerow([n, '未返却'])
            else:
                writer.writerow([n, d])

def main():
    start = time.time()
    results = parse_docx_tables("1-20.docx")
    write_results_to_csv(results, './抽出結果.txt')
    end = time.time()
    print(f"実行時間(秒):{end - start}")

if __name__ == "__main__":
    main()