"""
Programme permettant d'extraire des tableaux d'un PDF de trois manières différentes :
- Extraire tous les tableaux 
- Extraire les tableaux d'une ou plusieurs pages spécifiques
- Extraire le tableau d'une zone spécifique d'une page particulière

Problèmes actuels :
- La zone encadrée ne correspond pas à la zone extraite ensuite dans Excel (voir les fonctions "extract_tables_from_regions" et "update_manual_selection_canvas" pour corriger l'erreur).
- Les tableaux sont bien extraits, mais certains éléments d'une même cellule sont séparés.
- J'ai tenté de structurer ce code selon le modèle MVC, mais certains modes d'extraction ne fonctionnaient plus lors des tests. C'est pourquoi je vous envoie le code sur une seule page, qui fonctionne correctement.
"""





import dash
from dash import dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
import base64
import io
import pandas as pd
import fitz  # PyMuPDF
from flask import Flask, send_file
import plotly.graph_objs as go
from PIL import Image
import logging
import tabula
import pytesseract
from pdf2image import convert_from_bytes

# Create a Flask server instance
server = Flask(__name__)
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP], server=server)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Define global variables
excel_io = io.BytesIO()
pdf_images = {}  # To store base64 encoded images of PDF pages

app.layout = html.Div([
    dbc.Container([
        dbc.Row([
            dbc.Col(html.H1("PDF Table Extractor", className="text-center")),
        ]),
        dbc.Row([
            dbc.Col(dcc.Upload(
                id='upload-pdf',
                children=html.Div(['Drag and Drop or ', html.A('Select a PDF')]),
                style={
                    'width': '100%',
                    'height': '60px',
                    'lineHeight': '60px',
                    'borderWidth': '1px',
                    'borderStyle': 'dashed',
                    'borderRadius': '5px',
                    'textAlign': 'center',
                    'margin': '10px'
                },
                multiple=False
            )),
        ]),
        dbc.Row([
            dbc.Col(html.Div(id='pdf-preview')),
        ]),
        dbc.Row([
            dbc.Col([
                html.Label("Select Extraction Mode:"),
                dcc.RadioItems(
                    id='extraction-mode',
                    options=[
                        {'label': 'Extract All Tables', 'value': 'all'},
                        {'label': 'Select Pages', 'value': 'pages'},
                        {'label': 'Manual Selection', 'value': 'manual'}
                    ],
                    value='all'
                ),
                html.Br(),
                dcc.Input(
                    id='page-selection',
                    type='text',
                    placeholder='Enter page numbers (e.g., 1,3,5-7)',
                    style={'display': 'none'}
                ),
                html.Div(id='manual-selection-ui', style={'display': 'none'}),
                dcc.Dropdown(
                    id='manual-page-selection',
                    options=[],  # Will be updated dynamically
                    value=None,
                    style={'display': 'none'}
                ),
                dcc.Graph(
                    id='manual-selection-canvas',
                    config={
                        'scrollZoom': True,
                        'displayModeBar': True,
                    },
                    style={'height': '600px', 'display': 'none'}
                ),
            ]),
        ]),
        dbc.Row([
            dbc.Col(html.Button("Extract Tables", id="extract-button", className="btn btn-primary")),
        ]),
        dbc.Row([
            dbc.Col(html.Div(id='extracted-tables', className="mt-4")),
        ])
    ])
])

@app.callback(
    Output('pdf-preview', 'children'),
    Input('upload-pdf', 'contents')
)
def preview_pdf(pdf_content):
    """
    Callback to preview the uploaded PDF by converting its pages to images.

    Args:
        pdf_content (str): Base64 encoded content of the uploaded PDF.

    Returns:
        html.Div: HTML content displaying the preview of PDF pages as images.
    """
    if pdf_content is None:
        return html.Div("Upload a PDF to see the preview.")

    content_type, content_string = pdf_content.split(',')
    decoded = base64.b64decode(content_string)
    pdf_file = io.BytesIO(decoded)

    # Extract pages and convert to images
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    images = []
    global pdf_images
    pdf_images = {}
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap()
        img_bytes = io.BytesIO(pix.tobytes())
        img_base64 = base64.b64encode(img_bytes.getvalue()).decode('utf-8')
        pdf_images[page_num] = img_base64
        images.append(html.Img(src=f'data:image/png;base64,{img_base64}', style={'width': '100%', 'margin-bottom': '10px'}, id=f'page-{page_num}'))

    return html.Div(images)

@app.callback(
    Output('page-selection', 'style'),
    Output('manual-selection-ui', 'style'),
    Output('manual-page-selection', 'style'),
    Output('manual-selection-canvas', 'style'),
    Input('extraction-mode', 'value')
)
def update_ui_based_on_mode(extraction_mode):
    """
    Update the UI based on the selected extraction mode.

    Args:
        extraction_mode (str): Selected extraction mode (all, pages, manual).

    Returns:
        tuple: Visibility styles for different UI elements.
    """
    if extraction_mode == 'pages':
        return {'display': 'block'}, {'display': 'none'}, {'display': 'none'}, {'display': 'none'}
    elif extraction_mode == 'manual':
        return {'display': 'none'}, {'display': 'block'}, {'display': 'block'}, {'display': 'block'}
    return {'display': 'none'}, {'display': 'none'}, {'display': 'none'}, {'display': 'none'}

@app.callback(
    Output('manual-page-selection', 'options'),
    Input('pdf-preview', 'children')
)
def update_manual_page_selection_options(pdf_preview):
    """
    Dynamically update the dropdown options for manual page selection.

    Args:
        pdf_preview (html.Div): HTML content displaying the PDF preview.

    Returns:
        list: List of page options for the dropdown.
    """
    if pdf_preview:
        return [{'label': f'Page {i+1}', 'value': i} for i in range(len(pdf_images))]
    return []

@app.callback(
    Output('manual-selection-canvas', 'figure'),
    Input('manual-page-selection', 'value')
)
def update_manual_selection_canvas(selected_page):
    """
    Update the manual selection canvas with the selected page.

    Args:
        selected_page (int): The selected page number.

    Returns:
        go.Figure: Plotly figure with the selected PDF page as background.
    """
    if selected_page is None:
        return {}

    # Afficher la page sélectionnée en arrière-plan dans le graphique
    img_data = pdf_images[selected_page]
    img = Image.open(io.BytesIO(base64.b64decode(img_data)))
    img_width, img_height = img.size

    fig = go.Figure()

    # Définir l'image de fond
    fig.add_layout_image(
        dict(
            source=f'data:image/png;base64,{img_data}',
            xref="x",
            yref="y",
            x=0,
            y=img_height,
            sizex=img_width,
            sizey=img_height,
            sizing="stretch",
            opacity=1,
            layer="below"
        )
    )

    # Configurer les axes pour correspondre à la taille de l'image
    fig.update_xaxes(visible=False, range=[0, img_width])
    fig.update_yaxes(visible=False, range=[0, img_height])

    fig.update_layout(
        width=img_width,
        height=img_height,
        margin=dict(l=0, r=0, t=0, b=0),
        paper_bgcolor="white",
        plot_bgcolor="white",
        dragmode="drawrect",  # Activer le mode de dessin de rectangles
    )

    return fig

@app.callback(
    Output('extracted-tables', 'children'),
    Input('extract-button', 'n_clicks'),
    State('upload-pdf', 'contents'),
    State('extraction-mode', 'value'),
    State('page-selection', 'value'),
    State('manual-selection-canvas', 'relayoutData'),
    State('manual-page-selection', 'value')
)
def extract_tables(n_clicks, pdf_content, extraction_mode, pages, relayout_data, selected_page):
    """
    Extract tables from the uploaded PDF based on the selected extraction mode.

    Args:
        n_clicks (int): Number of clicks on the extract button.
        pdf_content (str): Base64 encoded content of the uploaded PDF.
        extraction_mode (str): Selected extraction mode.
        pages (str): Selected page numbers.
        relayout_data (dict): Data from the manual selection canvas.
        selected_page (int): Selected page number for manual extraction.

    Returns:
        html.Div: HTML content displaying the extracted tables.
    """
    if n_clicks is None or pdf_content is None:
        return html.Div("No PDF uploaded or no extraction requested.")

    if extraction_mode == 'manual' and (not relayout_data or selected_page is None):
        return html.Div("Please select a region on the PDF to extract tables from.")

    content_type, content_string = pdf_content.split(',')
    decoded = base64.b64decode(content_string)
    pdf_file = io.BytesIO(decoded)

    if extraction_mode == 'all':
        dfs = tabula.read_pdf(pdf_file, pages='all', multiple_tables=True, lattice=False)
    elif extraction_mode == 'pages':
        if pages:
            page_numbers = [int(p) for p in pages.replace(" ", "").split(',') if p.isdigit()]
            dfs = tabula.read_pdf(pdf_file, pages=page_numbers, multiple_tables=True, lattice=False)
        else:
            return html.Div("No pages selected.")
    elif extraction_mode == 'manual' and relayout_data:
        selected_regions = parse_relayout_data(relayout_data)
        dfs = extract_tables_from_regions(pdf_file, selected_page, selected_regions)
    else:
        dfs = []

    if dfs:
        global excel_io
        excel_io = io.BytesIO()
        with pd.ExcelWriter(excel_io, engine='openpyxl') as writer:
            for i, df in enumerate(dfs):
                df.to_excel(writer, sheet_name=f'Table_{i}', index=False)
        excel_io.seek(0)

        return html.A('Download Excel', href='/download/extracted-tables.xlsx')

    return html.Div("No tables found.")

def parse_relayout_data(relayout_data):
    """
    Parse the relayout data from the manual selection canvas.

    Args:
        relayout_data (dict): Data from the Plotly relayout event.

    Returns:
        list: List of tuples representing the selected regions (x0, y0, x1, y1).
    """
    # Fonction pour analyser les données de relayout pour extraire les régions sélectionnées
    if 'shapes' in relayout_data:
        return [(shape['x0'], shape['y0'], shape['x1'], shape['y1']) for shape in relayout_data['shapes']]
    return []

def extract_tables_from_regions(pdf_file, page_number, regions):
    """
    Extract tables from selected regions of a specific page in a PDF file.
    
    Args:
        pdf_file (io.BytesIO): The PDF file as a BytesIO object.
        page_number (int): The page number from which to extract tables.
        regions (list of tuples): List of tuples where each tuple represents a region 
                                  (x0, y0, x1, y1) for cropping.
    
    Returns:
        list of pd.DataFrame: List of DataFrames containing the extracted tables.
    """
    dfs = []
    images = convert_from_bytes(pdf_file.read(), first_page=page_number + 1, last_page=page_number + 1)
    page_image = images[0]

    img_width, img_height = page_image.size
    for i, region in enumerate(regions):
        x0, y0, x1, y1 = region
        
        # Ensure coordinates are in the correct order
        if x0 > x1:
            x0, x1 = x1, x0
        if y0 > y1:
            y0, y1 = y1, y0
        
        # Ensure coordinates are within the image bounds
        x0 = max(0, x0)
        y0 = max(0, y0)
        x1 = min(page_image.width, x1)
        y1 = min(page_image.height, y1)
        
        # Crop the image
        cropped_img = page_image.crop((x0, y0, x1, y1))
        
        # Log the cropped image
        logger.info(f"Cropped image {i + 1}: x0={x0}, y0={y0}, x1={x1}, y1={y1}")

        # Optionally save or display the cropped image
        cropped_img.save(f'cropped_image_{i + 1}.png')  # Save to a file
        cropped_img.show()  # Uncomment to display the image if running locally

        # Extract text using OCR
        text = pytesseract.image_to_string(cropped_img)
        
        # Convert the OCR text to a DataFrame
        data = [line.split() for line in text.split('\n') if line.strip()]
        df = pd.DataFrame(data)
        dfs.append(df)
    
    return dfs


@app.server.route('/download/extracted-tables.xlsx')
def download_tables():
    global excel_io
    if excel_io and excel_io.getvalue():
        # Reset the pointer to the beginning of the stream before sending
        excel_io.seek(0)

        return send_file(
            io.BytesIO(excel_io.read()),  # Create a new BytesIO object to avoid closing the original
            as_attachment=True,
            download_name='extracted-tables.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    return "No file to download", 404

if __name__ == '__main__':
    app.run_server(debug=True)

