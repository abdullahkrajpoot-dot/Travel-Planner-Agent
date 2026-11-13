# AI Travel Planner Agent

A professional Streamlit app for generating travel itineraries and downloadable PDF travel plans.

## Overview

This application helps travel agents create clean, day-by-day travel plans for clients. It generates itinerary titles, daily activities, destination places, image-based PDF pages, and a polished downloadable travel plan.

## Key Features

- **AI Travel Itinerary:** Creates day-wise travel plans from client, destination, and travel dates.
- **Arrival & Departure Rules:** First day is arrival/rest only, and last day is departure only with no visits.
- **Middle-Day Sightseeing:** Sightseeing places and images are used only for the main travel days.
- **PDF Download:** Generates a professional PDF travel plan with optimized image handling for smaller file size.
- **Streamlit Interface:** Simple local web app for entering trip details and downloading the final plan.
- **Prompt Revisions:** Users can describe a requested change and apply it to the generated itinerary.
- **Inline Document Editing:** Travel-plan text can be edited directly in the document preview.
- **PDF and Word Export:** The same updated itinerary can be downloaded as PDF or editable DOCX.
- **Rating:** A compact five-star feedback control is shown after the result.

## Working Structure

1. The user enters the client name, destination, title, and travel date range.
2. The app generates a structured day-by-day itinerary. If `OPENROUTER_API_KEY` is available it uses the configured language model; otherwise it uses the built-in itinerary generator.
3. Destination and landmark images are collected in parallel and optimized for document output.
4. The itinerary and images are stored in Streamlit session state as one `trip_result` object.
5. The user can revise the plan in two ways:
   - Enter a natural-language instruction in the prompt revision field.
   - Open Document Preview and edit the text directly in the rich-text editor.
6. Saving a change rebuilds both output formats from the same current `trip_result` data:
   - Professional PDF through `fpdf2`.
   - Editable Word document through `python-docx`.
7. The Download tab provides both files, and the final rating control records the user's score for the current session.

### Main Components

- `get_ai_plan`: generates the initial itinerary.
- `revise_plan_with_prompt`: applies a natural-language change request.
- `build_image_bundle`: fetches and prepares destination images.
- `st_quill`: provides click-to-edit rich-text fields in Document Preview.
- `build_pdf`: produces the PDF version.
- `build_docx`: produces the editable Word version.
- `rebuild_result_documents`: keeps PDF and Word output synchronized after edits.

## Tech Stack

- **Language:** Python 3.x
- **Framework:** Streamlit
- **PDF:** fpdf2
- **Images:** Pillow
- **Web Requests:** requests

## Installation & Setup

1. Create and activate a virtual environment:

   ```bash
   python -m venv .venv
   ```

   Windows:

   ```bash
   .\.venv\Scripts\activate
   ```

   Mac/Linux:

   ```bash
   source .venv/bin/activate
   ```

2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Run the application:

   ```bash
   streamlit run app.py
   ```

4. Open the local app:

   ```text
   http://127.0.0.1:8501
   ```

## Local Helper Scripts

- `start_localhost.bat` starts the Streamlit app on localhost.
- `start_localhost_hidden.vbs` starts the app in a hidden window.
- `stop_localhost_hidden.vbs` stops the hidden localhost server.
