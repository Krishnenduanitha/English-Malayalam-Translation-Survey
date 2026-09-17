# ============================================================
# ENGLISH → MALAYALAM TRANSLATION STUDY
# EMPHASIS ONLY
# ============================================================

import streamlit as st
import pandas as pd
import os
import random
import hashlib
import html
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="English-to-Malayalam Translation Study",
    page_icon="🎧",
    layout="centered"
)


# ============================================================
# PATHS
# ============================================================

BASE_DIR = os.path.dirname(
    os.path.abspath(__file__)
)

EXCEL_FILE = os.path.join(
    BASE_DIR,
    "25 sentences malayalam final.xlsx"
)

AUDIO_DIR = os.path.join(
    BASE_DIR,
    "audio"
)


# ============================================================
# GOOGLE SHEETS HEADERS
# ============================================================

HEADERS = [
    "participant_name",
    "age_range",
    "native_language",
    "english_proficiency",
    "malayalam_proficiency",
    "headphones",
    "hearing_difficulties",
    "speech_experience",
    "prosody_understanding",
    "listening_test_experience",
    "question_number",
    "english_sentence",
    "emphasized_word",
    "audiofile",
    "selected_translation",
    "selected_translation_type",
    "prosodic_rating",
    "response_time_seconds",
    "remarks",
    "last_updated"
]


# ============================================================
# SESSION STATE
# ============================================================

if "page" not in st.session_state:
    st.session_state.page = "welcome"

if "participant_name" not in st.session_state:
    st.session_state.participant_name = ""

if "demographics" not in st.session_state:
    st.session_state.demographics = {}

if "current_question" not in st.session_state:
    st.session_state.current_question = 0

if "answers" not in st.session_state:
    st.session_state.answers = {}

if "rating_selections" not in st.session_state:
    st.session_state.rating_selections = {}

if "translation_selections" not in st.session_state:
    st.session_state.translation_selections = {}

if "randomized_options" not in st.session_state:
    st.session_state.randomized_options = {}

if "question_start_times" not in st.session_state:
    st.session_state.question_start_times = {}

if "remarks" not in st.session_state:
    st.session_state.remarks = ""


# ============================================================
# GOOGLE SHEETS CONNECTION
# ============================================================

@st.cache_resource
def get_google_sheet():

    try:

        config = st.secrets[
            "connections"
        ][
            "gsheets"
        ]

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]

        credentials_info = {
            "type":
                config["type"],

            "project_id":
                config["project_id"],

            "private_key_id":
                config["private_key_id"],

            "private_key":
                config["private_key"],

            "client_email":
                config["client_email"],

            "client_id":
                config["client_id"],

            "auth_uri":
                config["auth_uri"],

            "token_uri":
                config["token_uri"],

            "auth_provider_x509_cert_url":
                config[
                    "auth_provider_x509_cert_url"
                ],

            "client_x509_cert_url":
                config[
                    "client_x509_cert_url"
                ]
        }

        credentials = (
            Credentials
            .from_service_account_info(
                credentials_info,
                scopes=scopes
            )
        )

        client = gspread.authorize(
            credentials
        )

        spreadsheet = client.open_by_url(
            config["spreadsheet"]
        )

        worksheet_name = config.get(
            "worksheet",
            "Responses"
        )

        worksheet = spreadsheet.worksheet(
            worksheet_name
        )

        return worksheet

    except Exception as e:

        st.error(
            "Could not connect to Google Sheets."
        )

        st.exception(e)

        return None


# ============================================================
# INITIALIZE GOOGLE SHEET
# ============================================================

def initialize_sheet():

    worksheet = get_google_sheet()

    if worksheet is None:
        return None

    try:

        values = worksheet.get_all_values()

        # Empty sheet
        if not values:

            worksheet.update(
                "A1:U1",
                [HEADERS],
                value_input_option="USER_ENTERED"
            )

            return worksheet

        existing_headers = [
            str(x).strip()
            for x in values[0]
        ]

        # Make header row exactly match our headers
        if existing_headers != HEADERS:

            worksheet.update(
                "A1:U1",
                [HEADERS],
                value_input_option="USER_ENTERED"
            )

        return worksheet

    except Exception as e:

        st.error(
            "Could not initialize Google Sheets."
        )

        st.exception(e)

        return None


# ============================================================
# READ GOOGLE SHEET
# ============================================================

def read_sheet():

    worksheet = initialize_sheet()

    if worksheet is None:
        return []

    try:

        return worksheet.get_all_values()

    except Exception as e:

        st.error(
            "Could not read Google Sheets."
        )

        st.exception(e)

        return []


# ============================================================
# CONVERT GOOGLE SHEET ROW TO DICTIONARY
# ============================================================

def row_to_dict(row):

    result = {}

    for index, header in enumerate(
        HEADERS
    ):

        if index < len(row):

            result[header] = str(
                row[index]
            ).strip()

        else:

            result[header] = ""

    return result


# ============================================================
# CHECK IF PARTICIPANT EXISTS
# ============================================================

def participant_exists(
    participant_name
):

    values = read_sheet()

    if len(values) <= 1:
        return False

    target = (
        participant_name
        .strip()
        .lower()
    )

    for row in values[1:]:

        data = row_to_dict(row)

        existing_name = (
            data["participant_name"]
            .strip()
            .lower()
        )

        if existing_name == target:

            return True

    return False


# ============================================================
# LOAD PARTICIPANT PROGRESS
# ============================================================

def load_participant_progress(
    participant_name
):

    values = read_sheet()

    answers = {}
    remarks = ""

    if len(values) <= 1:

        return {
            "answers": answers,
            "remarks": remarks,
            "first_unanswered": 0
        }

    target = (
        participant_name
        .strip()
        .lower()
    )

    for row in values[1:]:

        data = row_to_dict(row)

        existing_name = (
            data["participant_name"]
            .strip()
            .lower()
        )

        if existing_name != target:
            continue

        question_number = (
            data["question_number"]
            .strip()
        )

        if question_number:

            answers[
                question_number
            ] = {

                "selected_translation":
                    data[
                        "selected_translation"
                    ],

                "selected_translation_type":
                    data[
                        "selected_translation_type"
                    ],

                "prosodic_rating":
                    data[
                        "prosodic_rating"
                    ],

                "response_time_seconds":
                    data[
                        "response_time_seconds"
                    ]
            }

        if data["remarks"]:

            remarks = data[
                "remarks"
            ]

    # --------------------------------------------------------
    # Find first unanswered question
    # --------------------------------------------------------

    questions = load_questions()

    first_unanswered = len(
        questions
    )

    for index, row in questions.iterrows():

        question_number = str(
            index + 1
        )

        if question_number not in answers:

            first_unanswered = index

            break

    return {
        "answers": answers,
        "remarks": remarks,
        "first_unanswered": first_unanswered
    }


# ============================================================
# LOAD QUESTIONS FROM EXCEL
# ============================================================

@st.cache_data
def load_questions():

    if not os.path.exists(
        EXCEL_FILE
    ):

        st.error(
            "Excel file not found."
        )

        st.code(
            EXCEL_FILE
        )

        return pd.DataFrame()

    try:

        df = pd.read_excel(
            EXCEL_FILE
        )

    except Exception as e:

        st.error(
            "Could not read the Excel file."
        )

        st.exception(e)

        return pd.DataFrame()

    # --------------------------------------------------------
    # Clean column names
    # --------------------------------------------------------

    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    # --------------------------------------------------------
    # Required columns
    # --------------------------------------------------------

    required_columns = [

        "English Sentence",

        "Malayalam Sentence",

        "Google Translate",

        "Particles removed",

        "English Word",

        "Audio filename"
    ]

    missing_columns = [

        column

        for column in required_columns

        if column not in df.columns
    ]

    if missing_columns:

        st.error(
            "The following required columns "
            "are missing from the Excel file:"
        )

        st.write(
            missing_columns
        )

        st.write(
            "Columns found in the Excel file:"
        )

        st.write(
            list(df.columns)
        )

        return pd.DataFrame()

    # --------------------------------------------------------
    # Remove completely empty rows
    # --------------------------------------------------------

    df = df[
        df["English Sentence"]
        .fillna("")
        .astype(str)
        .str.strip()
        != ""
    ].copy()

    # --------------------------------------------------------
    # Remove accidental duplicate header row
    # --------------------------------------------------------

    header_mask = (

        df["English Sentence"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
        ==
        "english sentence"
    )

    df = df[
        ~header_mask
    ].copy()

    # --------------------------------------------------------
    # Reset index
    # --------------------------------------------------------

    df = df.reset_index(
        drop=True
    )

    # --------------------------------------------------------
    # Convert missing values to empty strings
    # --------------------------------------------------------

    df = df.fillna("")

    return df


questions_df = load_questions()


# ============================================================
# AUDIO PATH
# ============================================================

def get_audio_path(
    audio_filename
):

    audio_filename = str(
        audio_filename
    ).strip()

    if not audio_filename:

        return None

    # --------------------------------------------------------
    # Excel should contain only the filename.
    #
    # Example:
    #
    # 4xKgo0_HGRMThe_hidden_ways...wav
    #
    # --------------------------------------------------------

    filename = os.path.basename(
        audio_filename
    )

    audio_path = os.path.join(
        AUDIO_DIR,
        filename
    )

    if os.path.isfile(
        audio_path
    ):

        return audio_path

    # --------------------------------------------------------
    # Case-insensitive search
    # --------------------------------------------------------

    if os.path.exists(
        AUDIO_DIR
    ):

        target = filename.lower()

        for root, dirs, files in os.walk(
            AUDIO_DIR
        ):

            for file in files:

                if file.lower() == target:

                    return os.path.join(
                        root,
                        file
                    )

    return None


# ============================================================
# HIGHLIGHT EMPHASIZED WORD
# ============================================================

def highlight_emphasis(
    sentence,
    emphasized_word
):

    sentence = str(
        sentence
    )

    emphasized_word = str(
        emphasized_word
    ).strip()

    escaped_sentence = html.escape(
        sentence
    )

    if not emphasized_word:

        return escaped_sentence

    # Support comma-separated words
    words = [
        word.strip()
        for word in emphasized_word.split(",")
        if word.strip()
    ]

    result = escaped_sentence

    for word in words:

        escaped_word = html.escape(
            word
        )

        result = result.replace(
            escaped_word,
            (
                '<span class="emphasis-word">'
                + escaped_word
                + "</span>"
            )
        )

    return result


# ============================================================
# RANDOMIZE THREE TRANSLATION OPTIONS
# ============================================================

def get_randomized_options(
    participant_name,
    question_number,
    malayalam_sentence,
    google_translate,
    particles_removed
):

    options = [

        {
            "text":
                str(
                    malayalam_sentence
                ).strip(),

            "type":
                "human translation"
        },

        {
            "text":
                str(
                    google_translate
                ).strip(),

            "type":
                "machine translation"
        },

        {
            "text":
                str(
                    particles_removed
                ).strip(),

            "type":
                "particles removed"
        }
    ]

    # --------------------------------------------------------
    # Stable randomization
    #
    # Same participant + same question
    # = same option order.
    # --------------------------------------------------------

    seed_string = (
        participant_name.strip().lower()
        + "_"
        + str(question_number)
    )

    seed_hash = hashlib.sha256(
        seed_string.encode(
            "utf-8"
        )
    ).hexdigest()

    seed = int(
        seed_hash[:16],
        16
    )

    rng = random.Random(
        seed
    )

    rng.shuffle(
        options
    )

    return options


# ============================================================
# SAVE RESPONSE
# ============================================================

def save_response(
    row,
    question_number,
    selected_translation,
    selected_type,
    selected_rating,
    response_time,
    audio_filename
):

    worksheet = initialize_sheet()

    if worksheet is None:

        return False

    try:

        demographics = (
            st.session_state
            .demographics
        )

        data = {

            "participant_name":
                st.session_state
                .participant_name,

            "age_range":
                demographics.get(
                    "age_range",
                    ""
                ),

            "native_language":
                demographics.get(
                    "native_language",
                    ""
                ),

            "english_proficiency":
                demographics.get(
                    "english_proficiency",
                    ""
                ),

            "malayalam_proficiency":
                demographics.get(
                    "malayalam_proficiency",
                    ""
                ),

            "headphones":
                demographics.get(
                    "headphones",
                    ""
                ),

            "hearing_difficulties":
                demographics.get(
                    "hearing_difficulties",
                    ""
                ),

            "speech_experience":
                demographics.get(
                    "speech_experience",
                    ""
                ),

            "prosody_understanding":
                demographics.get(
                    "prosody_understanding",
                    ""
                ),

            "listening_test_experience":
                demographics.get(
                    "listening_test_experience",
                    ""
                ),

            "question_number":
                str(
                    question_number
                ),

            "english_sentence":
                str(
                    row["English Sentence"]
                ).strip(),

            "emphasized_word":
                str(
                    row["English Word"]
                ).strip(),

            "audiofile":
                audio_filename,

            "selected_translation":
                selected_translation,

            "selected_translation_type":
                selected_type,

            "prosodic_rating":
                selected_rating,

            "response_time_seconds":
                response_time,

            "remarks":
                "",

            "last_updated":
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
        }

        row_values = [

            data.get(
                header,
                ""
            )

            for header in HEADERS
        ]

        # ----------------------------------------------------
        # READ SHEET SAFELY
        # ----------------------------------------------------

        values = worksheet.get_all_values()

        existing_row_number = None

        participant_target = (
            st.session_state
            .participant_name
            .strip()
            .lower()
        )

        question_target = str(
            question_number
        ).strip()

        # ----------------------------------------------------
        # Find existing response
        # ----------------------------------------------------

        for row_number, existing_row in enumerate(
            values[1:],
            start=2
        ):

            existing_data = row_to_dict(
                existing_row
            )

            existing_participant = (
                existing_data[
                    "participant_name"
                ]
                .strip()
                .lower()
            )

            existing_question = (
                existing_data[
                    "question_number"
                ]
                .strip()
            )

            if (
                existing_participant
                ==
                participant_target
                and
                existing_question
                ==
                question_target
            ):

                existing_row_number = (
                    row_number
                )

                break

        # ----------------------------------------------------
        # Update
        # ----------------------------------------------------

        if existing_row_number:

            worksheet.update(
                f"A{existing_row_number}:U{existing_row_number}",
                [row_values],
                value_input_option="USER_ENTERED"
            )

        # ----------------------------------------------------
        # Append
        # ----------------------------------------------------

        else:

            worksheet.append_row(
                row_values,
                value_input_option="USER_ENTERED"
            )

        return True

    except Exception as e:

        st.error(
            "Could not save the response."
        )

        st.exception(e)

        return False


# ============================================================
# SAVE REMARKS
# ============================================================

def save_remarks():

    worksheet = initialize_sheet()

    if worksheet is None:

        return False

    try:

        values = worksheet.get_all_values()

        participant_target = (
            st.session_state
            .participant_name
            .strip()
            .lower()
        )

        found = False

        for row_number, row in enumerate(
            values[1:],
            start=2
        ):

            data = row_to_dict(
                row
            )

            existing_name = (
                data["participant_name"]
                .strip()
                .lower()
            )

            if (
                existing_name
                ==
                participant_target
            ):

                # T? No.
                # With our 21-column HEADERS:
                #
                # S = remarks? Let's calculate:
                #
                # 19 = response_time_seconds
                # 20 = remarks
                # 21 = last_updated
                #
                # Therefore:
                # T = remarks
                # U = last_updated

                worksheet.update_cell(
                    row_number,
                    20,
                    st.session_state.remarks
                )

                worksheet.update_cell(
                    row_number,
                    21,
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                )

                found = True

        return found

    except Exception as e:

        st.error(
            "Could not save remarks."
        )

        st.exception(e)

        return False


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
<style>

/* ============================================================
   MAIN TITLE
   ============================================================ */

.main-title {

    font-size: 40px;

    font-weight: 700;

    text-align: center;

    margin-top: 20px;

    margin-bottom: 10px;
}


.subtitle {

    text-align: center;

    font-size: 19px;

    margin-bottom: 35px;

}


/* ============================================================
   SECTION TITLE
   ============================================================ */

.section-title {

    font-size: 28px;

    font-weight: 700;

    margin-top: 25px;

    margin-bottom: 15px;
}


/* ============================================================
   INFO BOX
   ============================================================ */

.info-box {

    padding: 20px;

    border-radius: 12px;

    border: 1px solid
        rgba(120,120,120,0.4);

    margin-top: 20px;

    margin-bottom: 20px;
}


/* ============================================================
   ENGLISH SENTENCE
   ============================================================ */

.sentence-box {

    font-size: 22px;

    line-height: 1.8;

    padding: 25px;

    border-radius: 14px;

    border: 1px solid
        rgba(120,120,120,0.4);

    margin-top: 10px;

    margin-bottom: 20px;

    overflow-wrap: anywhere;
}


/* ============================================================
   EMPHASIZED WORD
   ============================================================ */

.emphasis-word {

    font-weight: 800;

    text-decoration: underline;

    text-decoration-thickness: 2px;
}


/* ============================================================
   TRANSLATION CARDS
   ============================================================ */

div[class*="st-key-translation_option_"] {

    width: 100% !important;

    margin: 0 !important;
}


div[class*="st-key-translation_option_"]
button {

    width: 100% !important;

    min-height: 125px !important;

    height: auto !important;

    padding: 25px 30px !important;

    margin: 0 !important;

    border-radius: 16px !important;

    border: 2px solid
        rgba(140,140,140,0.5) !important;

    text-align: left !important;

    white-space: normal !important;

    overflow: visible !important;

    overflow-wrap: anywhere !important;

    word-break: normal !important;

    box-sizing: border-box !important;

    transition:
        transform 0.15s ease,
        border-color 0.15s ease,
        box-shadow 0.15s ease !important;
}


div[class*="st-key-translation_option_"]
button:hover {

    transform: translateY(-2px) !important;

    border-color:
        rgba(77,163,255,0.9) !important;

    box-shadow:
        0 6px 18px
        rgba(77,163,255,0.18) !important;
}


div[class*="st-key-translation_option_"]
button p {

    font-size: 21px !important;

    line-height: 1.8 !important;

    white-space: normal !important;

    overflow-wrap: anywhere !important;

    word-break: normal !important;

    margin: 0 !important;

    padding: 0 !important;

    width: 100% !important;

    max-width: 100% !important;
}


/* ============================================================
   GAP BETWEEN OPTIONS
   ============================================================ */

.translation-gap {

    height: 30px;

    width: 100%;
}


/* ============================================================
   AUDIO
   ============================================================ */

audio {

    width: 100% !important;

    margin-top: 10px;

    margin-bottom: 20px;
}


/* ============================================================
   MOBILE
   ============================================================ */

@media (max-width: 768px) {

    .main-title {

        font-size: 31px;
    }

    .subtitle {

        font-size: 17px;
    }

    .section-title {

        font-size: 24px;
    }

    .sentence-box {

        font-size: 19px;

        padding: 18px;
    }

    div[class*="st-key-translation_option_"]
    button {

        min-height: 110px !important;

        padding: 20px 22px !important;
    }

    div[class*="st-key-translation_option_"]
    button p {

        font-size: 18px !important;

        line-height: 1.7 !important;
    }

    .translation-gap {

        height: 22px;
    }
}

</style>
""",
    unsafe_allow_html=True
)


# ============================================================
# WELCOME PAGE
# ============================================================

if st.session_state.page == "welcome":

    st.markdown(
        '<div class="main-title">'
        'English-to-Malayalam Translation Study'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitle">'
        'A study on the perception and transfer of '
        'prosodic emphasis from English speech '
        'into Malayalam translation'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        "## About the Study"
    )

    st.write(
        """
This study investigates how **prosodic emphasis** in
English speech is perceived and reflected in Malayalam
translation.

In spoken language, speakers may emphasize particular
words to convey focus, contrast, importance, or intended
meaning.

The purpose of this study is to examine how listeners
perceive emphasis in English speech and how this emphasis
is represented in different Malayalam versions.

During the study, you will listen to English speech
recordings, rate how strongly you perceive the emphasis,
and then choose the Malayalam version that best reflects
the intended meaning and emphasis of the original speech.

"""
    )

    st.markdown(
        """
<div class="info-box">

<b>Important:</b>

Please base your responses on your own perception of
the audio. There are no right or wrong answers from
the participant's perspective.

Note: Only focus on the given emphasized word there might be 
other emphasized words in the audio , just ignore those.

</div>
""",
        unsafe_allow_html=True
    )

    if st.button(
        "Begin Study →",
        type="primary",
        use_container_width=True
    ):

        st.session_state.page = (
            "participant_info"
        )

        st.rerun()


# ============================================================
# PARTICIPANT INFORMATION
# ============================================================

elif st.session_state.page == "participant_info":

    st.markdown(
        '<div class="section-title">'
        'Participant Information'
        '</div>',
        unsafe_allow_html=True
    )

    st.write(
        "Please provide the following information "
        "before beginning the study."
    )

    participant_name = st.text_input(
        "Participant name",
        value=st.session_state.participant_name,
        placeholder="Enter your name"
    )

    age_range = st.radio(
        "Age range",
        [
            "Below 18",
            "18–24",
            "25–34",
            "35–44",
            "45–54",
            "55 or above"
        ],
        index=None,
        key="participant_age"
    )

    native_language = st.text_input(
        "Native language(s)",
        key="participant_native_language"
    )

    english_proficiency = st.radio(
        "English proficiency",
        [
            "Beginner",
            "Intermediate",
            "Advanced",
            "Native / Near-native"
        ],
        index=None,
        key="participant_english"
    )

    malayalam_proficiency = st.radio(
        "Malayalam proficiency",
        [
            "None",
            "Beginner",
            "Intermediate",
            "Advanced",
            "Native"
        ],
        index=None,
        key="participant_malayalam"
    )

    headphones = st.radio(
        "Are you using headphones or earphones?",
        [
            "Yes",
            "No"
        ],
        index=None,
        key="participant_headphones"
    )

    hearing_difficulties = st.radio(
        "Do you have any difficulty hearing speech?",
        [
            "Yes",
            "No",
            "Prefer not to say"
        ],
        index=None,
        key="participant_hearing"
    )

    speech_experience = st.radio(
        "Do you have previous experience with "
        "speech, linguistics, audio, or related research?",
        [
            "Yes",
            "No"
        ],
        index=None,
        key="participant_speech"
    )

    prosody_understanding = st.radio(
        "How familiar are you with the concept of prosody?",
        [
            "Not familiar",
            "Slightly familiar",
            "Moderately familiar",
            "Very familiar"
        ],
        index=None,
        key="participant_prosody"
    )

    listening_test_experience = st.radio(
        "Have you participated in a listening test before?",
        [
            "Yes",
            "No"
        ],
        index=None,
        key="participant_listening"
    )

    st.markdown("---")

    if st.button(
        "Continue →",
        type="primary",
        use_container_width=True
    ):

        if not participant_name.strip():

            st.warning(
                "Please enter your participant name."
            )

        elif not age_range:

            st.warning(
                "Please select your age range."
            )

        elif not native_language.strip():

            st.warning(
                "Please enter your native language."
            )

        elif not english_proficiency:

            st.warning(
                "Please select your English proficiency."
            )

        elif not malayalam_proficiency:

            st.warning(
                "Please select your Malayalam proficiency."
            )

        elif not headphones:

            st.warning(
                "Please answer the headphones question."
            )

        elif not hearing_difficulties:

            st.warning(
                "Please answer the hearing question."
            )

        elif not speech_experience:

            st.warning(
                "Please answer the speech experience question."
            )

        elif not prosody_understanding:

            st.warning(
                "Please select your familiarity with prosody."
            )

        elif not listening_test_experience:

            st.warning(
                "Please answer the listening-test question."
            )

        else:

            st.session_state.participant_name = (
                participant_name.strip()
            )

            st.session_state.demographics = {

                "age_range":
                    age_range,

                "native_language":
                    native_language.strip(),

                "english_proficiency":
                    english_proficiency,

                "malayalam_proficiency":
                    malayalam_proficiency,

                "headphones":
                    headphones,

                "hearing_difficulties":
                    hearing_difficulties,

                "speech_experience":
                    speech_experience,

                "prosody_understanding":
                    prosody_understanding,

                "listening_test_experience":
                    listening_test_experience
            }

            # ------------------------------------------------
            # LOAD EXISTING PROGRESS
            # ------------------------------------------------

            progress = (
                load_participant_progress(
                    st.session_state
                    .participant_name
                )
            )

            if participant_exists(
                st.session_state
                .participant_name
            ):

                st.session_state.answers = (
                    progress["answers"]
                )

                st.session_state.remarks = (
                    progress["remarks"]
                )

                st.session_state.current_question = (
                    progress["first_unanswered"]
                )

                if (
                    st.session_state.current_question
                    >= len(questions_df)
                ):

                    st.session_state.page = (
                        "completed"
                    )

                else:

                    st.session_state.page = (
                        "instructions"
                    )

            else:

                st.session_state.current_question = 0

                st.session_state.page = (
                    "instructions"
                )

            st.rerun()


# ============================================================
# INSTRUCTIONS
# ============================================================

elif st.session_state.page == "instructions":

    st.markdown(
        '<div class="section-title">'
        'Instructions'
        '</div>',
        unsafe_allow_html=True
    )

    st.write(
        "Please read the instructions carefully."
    )

    st.markdown(
        """
### What you need to do

**1. Listen carefully to the English audio.**

You may replay the recording as many times as necessary.

**2. Pay attention to the emphasized word(s).**

The English sentence displayed on the screen contains
the word or words relevant to the emphasis being studied.

Listen carefully to how strongly the speaker emphasizes
these word(s).

**3. Rate the perceived emphasis.**

Before choosing a translation, rate how strongly you
perceived the indicated word(s) to be emphasized in
the English audio.

**4. Choose one Malayalam version.**

You will see three Malayalam versions of the sentence.

They represent:

- A Malayalam reference translation
- A machine-generated translation
- A version with the emphasis particle removed

The labels identifying these versions will not be shown
to you.

Choose the version that best reflects the intended meaning
and emphasis of the English speech.
"""
    )

    st.markdown(
        """
### Important points

- Listen carefully before answering.
- You may replay the audio as many times as necessary.
- Focus specifically on **emphasis** in the English speech.
- Give your emphasis rating **before** selecting a
  Malayalam version.
- Consider the intended meaning of the emphasized word.
- Base your answers on your own perception.
- Do not judge the speaker based on voice, gender,
  accent, speaking speed, or loudness.
- There are no right or wrong answers from the
  participant's perspective.
"""
    )

    st.markdown(
        """
### Emphasis rating scale

| Rating | Meaning |
|---|---|
| **1** | Not perceived at all |
| **2** | Slightly perceived |
| **3** | Moderately perceived |
| **4** | Strongly perceived |
| **5** | Very strongly perceived |
"""
    )

    if st.button(
        "Start Experiment →",
        type="primary",
        use_container_width=True
    ):

        st.session_state.page = (
            "experiment"
        )

        st.rerun()


# ============================================================
# EXPERIMENT
# ============================================================

elif st.session_state.page == "experiment":

    total_questions = len(
        questions_df
    )

    current_index = (
        st.session_state.current_question
    )

    # --------------------------------------------------------
    # Check completion
    # --------------------------------------------------------

    if current_index >= total_questions:

        st.session_state.page = (
            "remarks"
        )

        st.rerun()

    # --------------------------------------------------------
    # Current question
    # --------------------------------------------------------

    row = questions_df.iloc[
        current_index
    ]

    question_number = (
        current_index + 1
    )

    english_sentence = str(
        row["English Sentence"]
    ).strip()

    malayalam_sentence = str(
        row["Malayalam Sentence"]
    ).strip()

    google_translate = str(
        row["Google Translate"]
    ).strip()

    particles_removed = str(
        row["Particles removed"]
    ).strip()

    emphasized_word = str(
        row["English Word"]
    ).strip()

    audio_filename = str(
        row["Audio filename"]
    ).strip()

    # --------------------------------------------------------
    # Find EXACT audio from Excel filename
    # --------------------------------------------------------

    audio_path = get_audio_path(
        audio_filename
    )

    # ========================================================
    # PROGRESS
    # ========================================================

    st.progress(
        current_index / total_questions
    )

    st.write(
        f"Question {question_number} "
        f"of {total_questions}"
    )

    st.markdown("---")

    # ========================================================
    # ENGLISH SENTENCE
    # ========================================================

    st.markdown(
        '<div class="section-title">'
        'English sentence'
        '</div>',
        unsafe_allow_html=True
    )

    highlighted_sentence = (
        highlight_emphasis(
            english_sentence,
            emphasized_word
        )
    )

    st.markdown(
        f"""
<div class="sentence-box">
{highlighted_sentence}
</div>
""",
        unsafe_allow_html=True
    )

    if emphasized_word:

        st.write(
            f"**Indicated word(s):** "
            f"{emphasized_word}"
        )

    # ========================================================
    # AUDIO
    # ========================================================

    st.markdown(
        '<div class="section-title">'
        'Listen to the audio'
        '</div>',
        unsafe_allow_html=True
    )

    if audio_path:

        try:

            with open(
                audio_path,
                "rb"
            ) as audio_file:

                audio_bytes = (
                    audio_file.read()
                )

            st.audio(
                audio_bytes
            )

        except Exception as e:

            st.error(
                "The audio file could not be opened."
            )

            st.exception(e)

    else:

        st.error(
            "Audio file could not be found."
        )

        st.write(
            f"Expected filename:"
        )

        st.code(
            audio_filename
        )

        st.write(
            "Expected location:"
        )

        st.code(
            os.path.join(
                AUDIO_DIR,
                audio_filename
            )
        )

    # ========================================================
    # TIMER
    # ========================================================

    if question_number not in (
        st.session_state
        .question_start_times
    ):

        st.session_state.question_start_times[
            question_number
        ] = datetime.now()

    # ========================================================
    # EMPHASIS RATING
    # ========================================================

    st.markdown("---")

    st.markdown(
        '<div class="section-title">'
        'Rate the emphasis in the audio'
        '</div>',
        unsafe_allow_html=True
    )

    st.write(
        "How strongly did you perceive the "
        "indicated word(s) being emphasized "
        "in the audio?"
    )

    st.caption(
        "Please rate the audio before choosing "
        "a Malayalam translation."
    )

    rating_options = [

        "1 — Not perceived at all",

        "2 — Slightly perceived",

        "3 — Moderately perceived",

        "4 — Strongly perceived",

        "5 — Very strongly perceived"
    ]

    previous_rating = (
        st.session_state
        .rating_selections
        .get(
            question_number,
            None
        )
    )

    rating_index = None

    if previous_rating in rating_options:

        rating_index = (
            rating_options.index(
                previous_rating
            )
        )

    selected_rating = st.radio(

        "Emphasis rating",

        rating_options,

        index=rating_index,

        key=f"rating_{question_number}"
    )

    # ========================================================
    # THREE MALAYALAM OPTIONS
    # ========================================================

    st.markdown("---")

    st.markdown(
        '<div class="section-title">'
        'Choose the Malayalam version'
        '</div>',
        unsafe_allow_html=True
    )

    st.write(
        "Choose the version that best reflects "
        "the intended meaning and emphasis of "
        "the English speech."
    )

    # --------------------------------------------------------
    # Check all three options
    # --------------------------------------------------------

    if not malayalam_sentence:

        st.error(
            "Malayalam Sentence is missing "
            "for this question."
        )

        st.stop()

    if not google_translate:

        st.error(
            "Google Translate is missing "
            "for this question."
        )

        st.stop()

    if not particles_removed:

        st.error(
            "Particles removed is missing "
            "for this question."
        )

        st.stop()

    # --------------------------------------------------------
    # RANDOMIZE
    # --------------------------------------------------------

    options = get_randomized_options(

        st.session_state.participant_name,

        question_number,

        malayalam_sentence,

        google_translate,

        particles_removed
    )

    st.session_state.randomized_options[
        question_number
    ] = options

    # --------------------------------------------------------
    # Previous selection
    # --------------------------------------------------------

    previous_translation = (
        st.session_state
        .translation_selections
        .get(
            question_number,
            None
        )
    )

    # ========================================================
    # DISPLAY THREE SEPARATE CARDS
    # ========================================================

    for option_index, option in enumerate(
        options
    ):

        option_text = option[
            "text"
        ]

        is_selected = (
            previous_translation
            ==
            option_text
        )

        # ----------------------------------------------------
        # Selected option styling
        # ----------------------------------------------------

        if is_selected:

            st.markdown(
                f"""
<style>

div[class*="st-key-translation_option_{question_number}_{option_index}"]
button {{

    border:
        3px solid #4da3ff !important;

    box-shadow:
        0 0 0 1px #4da3ff,
        0 6px 20px
        rgba(77,163,255,0.22) !important;
}}

</style>
""",
                unsafe_allow_html=True
            )

        # ----------------------------------------------------
        # Button
        # ----------------------------------------------------

        if st.button(

            option_text,

            key=(
                f"translation_option_"
                f"{question_number}_"
                f"{option_index}"
            ),

            use_container_width=True
        ):

            st.session_state.translation_selections[
                question_number
            ] = option_text

            st.rerun()

        # ----------------------------------------------------
        # SPACE BETWEEN CARDS
        # ----------------------------------------------------

        if option_index < 2:

            st.markdown(
                '<div class="translation-gap"></div>',
                unsafe_allow_html=True
            )

    # --------------------------------------------------------
    # Current selected translation
    # --------------------------------------------------------

    selected_translation = (
        st.session_state
        .translation_selections
        .get(
            question_number,
            None
        )
    )

    # ========================================================
    # NAVIGATION
    # ========================================================

    st.markdown("---")

    col1, col2 = st.columns(
        2
    )

    # --------------------------------------------------------
    # Previous
    # --------------------------------------------------------

    with col1:

        if current_index > 0:

            if st.button(
                "← Previous",
                use_container_width=True
            ):

                st.session_state.current_question -= 1

                st.rerun()

    # --------------------------------------------------------
    # Next
    # --------------------------------------------------------

    with col2:

        if st.button(
            "Next →",
            type="primary",
            use_container_width=True
        ):

            # ------------------------------------------------
            # Rating required
            # ------------------------------------------------

            if not selected_rating:

                st.warning(
                    "Please provide an emphasis rating."
                )

                st.stop()

            # ------------------------------------------------
            # Translation required
            # ------------------------------------------------

            if not selected_translation:

                st.warning(
                    "Please select one Malayalam version."
                )

                st.stop()

            # ------------------------------------------------
            # Identify actual source internally
            # ------------------------------------------------

            selected_type = ""

            for option in options:

                if (
                    option["text"]
                    ==
                    selected_translation
                ):

                    selected_type = (
                        option["type"]
                    )

                    break

            # ------------------------------------------------
            # Response time
            # ------------------------------------------------

            start_time = (
                st.session_state
                .question_start_times
                .get(
                    question_number,
                    datetime.now()
                )
            )

            response_time = (

                datetime.now()
                -
                start_time

            ).total_seconds()

            response_time = round(
                response_time,
                2
            )

            # ------------------------------------------------
            # Store answer
            # ------------------------------------------------

            st.session_state.answers[
                str(question_number)
            ] = {

                "selected_translation":
                    selected_translation,

                "selected_translation_type":
                    selected_type,

                "prosodic_rating":
                    selected_rating,

                "response_time_seconds":
                    response_time
            }

            st.session_state.rating_selections[
                question_number
            ] = selected_rating

            st.session_state.translation_selections[
                question_number
            ] = selected_translation

            # ------------------------------------------------
            # Save
            # ------------------------------------------------

            success = save_response(

                row,

                question_number,

                selected_translation,

                selected_type,

                selected_rating,

                response_time,

                audio_filename
            )

            if success:

                # ------------------------------------------------
                # Move forward
                # ------------------------------------------------

                st.session_state.current_question += 1

                if (
                    st.session_state.current_question
                    >= total_questions
                ):

                    st.session_state.page = (
                        "remarks"
                    )

                st.rerun()


# ============================================================
# REMARKS PAGE
# ============================================================

elif st.session_state.page == "remarks":

    st.markdown(
        '<div class="section-title">'
        'Study Completed'
        '</div>',
        unsafe_allow_html=True
    )

    st.write(
        """
You have completed all the listening and
translation questions.

You may provide any comments or feedback
about your experience below.
"""
    )

    st.write(
        """
You may comment on:

- Clarity of the instructions
- Audio quality
- Difficulty in identifying emphasis
- Malayalam translation quality
- Ease of choosing between the three versions
- Your perception of emphasis
- Any technical issues
- Anything else you would like to mention
"""
    )

    remarks = st.text_area(

        "Additional comments or feedback (optional)",

        value=st.session_state.remarks,

        height=180,

        placeholder=(
            "Please share any comments "
            "or suggestions..."
        )
    )

    st.session_state.remarks = remarks

    if st.button(
        "Submit →",
        type="primary",
        use_container_width=True
    ):

        save_remarks()

        st.session_state.page = (
            "completed"
        )

        st.rerun()


# ============================================================
# COMPLETED PAGE
# ============================================================

elif st.session_state.page == "completed":

    st.markdown(
        '<div class="main-title">'
        'Thank You!'
        '</div>',
        unsafe_allow_html=True
    )

    st.success(
        "Your responses have been recorded successfully."
    )

    st.write(
        """
Thank you for participating in the
English-to-Malayalam translation study.

Your responses will help us understand how
prosodic emphasis in English speech is perceived
and reflected in Malayalam translation.

You may now close this page.
"""
    )