/*************************************************
 * Config.example.js
 *
 * Copy this file to Config.js and fill in your own IDs.
 * Config.js is listed in .gitignore and must never be committed.
 *************************************************/

const CONFIG = {
  MDM_SPREADSHEET_ID: "YOUR_SPREADSHEET_ID",
  MDM_EVENTS_SHEET: "Events",

  /*************************************************
   * SHEETS
   *************************************************/
  CONFIG_SHEET: "Config",
  CACHE_SHEET: "Vendor_Cache",
  EVENT_TEMPLATE_SHEET: "Event_Template",
  VENDOR_SEARCH: "Vendor_Search",
  VENDOR_UPDATE_QUEUE: "Vendor_Update_Queue",

  ROW: {
    VENDOR_SEARCH: {
      HEADER: 7,
      FIRST_DATA: 8
    }
  },

  SYSTEM_SHEETS: [
    "README",
    "Vendor_Cache",
    "Vendor_Search",
    "Email_Template",
    "Vendor_Update_Queue",
    "Config",
    "User_Guide",
    "VENDOR_UPDATE_QUEUE"
  ],

  /*************************************************
   * DRIVE
   *************************************************/
  RFQ_ROOT_FOLDER_ID: "YOUR_RFQ_ROOT_FOLDER_ID",
  DRIVE_FOLDER_ID: "YOUR_DRIVE_FOLDER_ID",
  RFQ_TEMPLATE_FOLDER_ID: "YOUR_RFQ_TEMPLATE_FOLDER_ID",
  FORM_TEMPLATE_ID: "YOUR_FORM_TEMPLATE_ID",
  EVALUATION_TEMPLATE_ID: "YOUR_EVALUATION_TEMPLATE_ID",

  /*************************************************
   * BLOCKS
   *************************************************/
  BLOCK: {
    START_MARKER: "MM:BLOCK_START",
    END_MARKER: "MM:BLOCK_END",
    SPACING: 1
  },

  /*************************************************
   * EVENT BLOCK TEMPLATES
   *************************************************/
  TEMPLATES: {
    Videography: "Videography_Block",
    Photography: "Photography_Block",
    Printing: "Printing_Block",
    Livestream: "Livestream_Block",
    Other: "Other_Block",
    AudioVisual: "AV_Block"
  },

  /*************************************************
   * EMAIL TEMPLATES (Google Doc IDs)
   *************************************************/
  EMAIL_TEMPLATES: {
    videography: "YOUR_VIDEOGRAPHY_TEMPLATE_DOC_ID",
    photography: "YOUR_PHOTOGRAPHY_TEMPLATE_DOC_ID",
    printing: "YOUR_PRINTING_TEMPLATE_DOC_ID",
    livestream: "YOUR_LIVESTREAM_TEMPLATE_DOC_ID",
    other: "YOUR_OTHER_TEMPLATE_DOC_ID",
    audiovisual: "YOUR_AUDIOVISUAL_TEMPLATE_DOC_ID"
  },

  /*************************************************
   * DYNAMIC PRICE COLUMNS
   *************************************************/
  DYNAMIC_PRICE_COLUMNS: [
    "printing",
    "other"
  ],

  /*************************************************
   * FORM → EVENT COLUMN MAPPING
   *************************************************/
  FORM_FIELDS: {
    COMPANY: "Company Name",
    TYPE: "What products or services do you provide?",
    EMAIL: "Email Address",
    WEBSITE: "Company Website URL",
    CONTACT: "Contact Name",
    SAMPLE: "Sample of Work (Google Drive / Portfolio URL)",
    QUOTATION: " (Google Drive URL)",
    FULL_DAY: "Full-day rate (USD)",
    HALF_DAY: "Half-day rate (USD)",
    CREW_NUMBER: "Crew Number",
    EQUIPMENT: "Equipment List",
    COMMENTS: "Please provide a brief breakdown of your pricing or additional comments"
  }
};

const EMAIL = {
  VENDOR: "vendor@example.com"
};

const FIELD_MAP = {

  ID: {
    event: "vendor_id"
  },

  NAME: {
    event: "vendor_name",
    cache: 2
  },

  TYPE: {
    event: "vendor_type",
    cache: 3
  },

  EMAIL: {
    event: "email",
    cache: 4
  },

  WEBSITE: {
    event: "website",
    cache: 5
  },

  CONTACT: {
    event: "contact_person",
    cache: 6
  },

  PORTFOLIO: {
    event: "Portfolio"
  },

  DECISION: {
    event: "Decision"
  },

  Equipment: {
    event: "Equipment List"
  },
  
  QUOTATION: {
    event: "Quotation (Raw)"
  }

};

const EVENT = {

  NAME: "C1",
  COUNTRY: "C2",
  CITY: "C3",
  DATE: "C4",
  PM: "C5",

  FOLDER_URL: "H1",
  EDIT_FORM_URL: "H2",
  PUBLIC_FORM_URL: "H3",
  DRAFT_EMAIL: "H4",
  EVALUATION_URL: "H5"

};