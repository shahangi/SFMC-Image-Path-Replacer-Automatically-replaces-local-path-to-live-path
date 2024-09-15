from flask import Flask, request, render_template, send_file, abort
import pandas as pd
from bs4 import BeautifulSoup
import os
import traceback
import re

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'html', 'xlsx'}

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def load_image_mapping(file_path):
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"Excel file not found: {file_path}")

    try:
        df = pd.read_excel(file_path, sheet_name='Sheet1')
        print("Columns found in the Excel file:", df.columns.tolist())

        local_image_col = next((col for col in df.columns if 'Local Image Name' in col.strip()), None)
        sfmc_url_col = next((col for col in df.columns if 'Live SFMC URL' in col.strip()), None)

        if local_image_col is None or sfmc_url_col is None:
            raise KeyError("Required columns not found in Excel file.")

        df[local_image_col] = df[local_image_col].astype(str)
        df[sfmc_url_col] = df[sfmc_url_col].astype(str)

        return dict(zip(df[local_image_col], df[sfmc_url_col]))

    except Exception as e:
        print(f"Error loading Excel file: {e}")
        raise

def normalize_text(text):
    """Normalize text by removing extra spaces and newlines."""
    return re.sub(r'\s+', ' ', text).strip()

def replace_text_in_html(content, from_text, to_text):
    """Replace all occurrences of `from_text` with `to_text` in the HTML content."""
    normalized_from_text = normalize_text(from_text)
    normalized_to_text = normalize_text(to_text)

    # Normalize HTML content
    normalized_content = normalize_text(content)

    print(f"Normalized 'from_text': '{normalized_from_text}'")
    print(f"Normalized 'to_text': '{normalized_to_text}'")

    # Perform direct string replacement
    updated_content = normalized_content.replace(normalized_from_text, normalized_to_text)

    # Reverse normalization to preserve original formatting as much as possible
    # Note: This step might require further adjustments depending on HTML complexity
    updated_content = re.sub(r'\s+', ' ', updated_content)

    # Debugging: Print snippets before and after replacement
    print(f"Content snippet before replacement: '{content[:500]}'")
    print(f"Content snippet after replacement: '{updated_content[:500]}'")

    return updated_content, f"Replaced '{from_text}' with '{to_text}'"

def process_files(html_file_path, excel_file_path, from_text, to_text, remove_preheader):
    mapping = load_image_mapping(excel_file_path)

    # Step 1: Read and print initial HTML content
    with open(html_file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    print(f"Original HTML content (first 1000 chars):\n{content[:1000]}")  # Debugging

    # Step 2: Handle text replacement
    salutation_message = "No change"
    if from_text:
        content, salutation_message = replace_text_in_html(content, from_text, to_text)

    # Step 3: Parse HTML with BeautifulSoup
    soup = BeautifulSoup(content, 'lxml')
    print(f"Parsed HTML content (first 1000 chars):\n{soup.prettify()[:1000]}")  # Debugging

    # Step 4: Handle image URL replacement
    not_found_images = []
    for img_tag in soup.find_all('img'):
        src = img_tag.get('src')
        if src:
            src = str(src)  # Ensure src is treated as a string
            matched = False
            for local_name, sfmc_url in mapping.items():
                local_name = str(local_name)  # Ensure local_name is treated as a string
                if local_name in src:
                    img_tag['src'] = sfmc_url
                    matched = True
                    break
            if not matched:
                not_found_images.append(src)

    # Step 5: Replace </html> with <custom name="opencounter" type="tracking"/>
    if '</html>' in str(soup):
        str_soup = str(soup)
        str_soup = str_soup.replace('</html>', '<custom name="opencounter" type="tracking"/>')
        soup = BeautifulSoup(str_soup, 'lxml')

    # Step 6: Remove <tr> containing {{customText}} if the checkbox is checked
    preheader_message = ""
    if remove_preheader:
        tables_to_check = soup.find_all('table', class_='main_body device_width')
        print(f"Tables to check for rows to remove (count: {len(tables_to_check)}):")

        for table in tables_to_check:
            rows_to_remove = []
            for row in table.find_all('tr'):
                if '{{customText' in row.get_text():
                    rows_to_remove.append(row)

            print(f"Rows to remove (count: {len(rows_to_remove)}):")
            for row in rows_to_remove:
                print(f"Row content:\n{row.prettify()}")  # Debugging

            for row in rows_to_remove:
                row.decompose()

        preheader_message = "Preheader removed successfully"

    # Step 7: Verify final HTML before writing to file
    final_html = str(soup)
    print(f"Final HTML content (first 1000 chars):\n{final_html[:1000]}")  # Debugging

    # Step 8: Write final HTML to file
    output_file = os.path.join(app.config['UPLOAD_FOLDER'], 'modified_' + os.path.basename(html_file_path))
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write(final_html)
    print(f"Modified HTML content written to {output_file}")  # Debugging

    return output_file, not_found_images, salutation_message, preheader_message

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        html_file = request.files.get('html_file')
        excel_file = request.files.get('excel_file')
        from_text = request.form.get('from_text', '').strip()
        to_text = request.form.get('to_text', '').strip()
        remove_preheader = 'remove_preheader' in request.form

        if html_file and allowed_file(html_file.filename) and excel_file and allowed_file(excel_file.filename):
            html_path = os.path.join(app.config['UPLOAD_FOLDER'], html_file.filename)
            excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)

            print(f"Saving HTML file to {html_path}")  # Debug print
            print(f"Saving Excel file to {excel_path}")  # Debug print

            html_file.save(html_path)
            excel_file.save(excel_path)

            try:
                output_file, not_found_images, salutation_message, preheader_message = process_files(html_path, excel_path, from_text, to_text, remove_preheader)
                output_filename = os.path.basename(output_file)
                return render_template('result.html', output_file=output_filename, not_found_images=not_found_images, salutation_message=salutation_message, preheader_message=preheader_message)
            except Exception as e:
                error_message = f"An error occurred: {e}\n{traceback.format_exc()}"
                print(error_message)  # Log to console
                return error_message, 500

    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(file_path):
        print(f"Serving file from {file_path}")  # Debug print
        return send_file(file_path, as_attachment=True)
    else:
        print(f"File not found at {file_path}")  # Debug print
        abort(404, description="File not found")

if __name__ == '__main__':
    app.run(debug=True)
