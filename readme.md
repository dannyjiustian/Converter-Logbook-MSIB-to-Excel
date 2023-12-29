
# Converter for Logbook MSIB Kampus Merdeka 

This code runs in NodeJS, it is hoped that this code can make it easier to convert existing logbooks in MSIB to excel format for campus or personal use.


## Installation

Download or Clone this repository. After that you open terminal in project and type command like bellow and enter, for install all package used. 👇 

**MAKE SURE YOUR COMPUTER HAS NODEJS AND MICROSOFT EXCEL INSTALLED TO OPEN EXCEL FILES.**

```bash
npm install
```
    
## Running the server

The code can running if you type and enter 👇 

```bash
npm start
```
## API MSIB Kampus Merdeka

#### Get data from Kampus Merdeka Platform

```http
  GET /https://api.kampusmerdeka.kemdikbud.go.id/magang/report/allweeks/xxxxxx
```

| Needs | Type     | Description             |
| :-------- | :------- | :-----------------------|
| `xxxxxx` | `number` | **Required**. Type of ID Kegiatan |
| `bearer token` | `string` | **Required**. Get From Kampus Merdeka after Login |

When the API runs successfully it will get the data as below 👇 
```json
{
  "data": [
    {
      "id": "xxxx",
      "position_id": "xxxx",
      "uuid_akun": "xxxx",
      "id_reg_penawaran": "xxxx",
      "start_date": "xxxx",
      "end_date": "xxxx",
      "counter": "xxxx",
      "status": "xxxx",
      "learned_weekly": "xxxx",
      "note_from_mentor": "xxxx",
      "reviewed_by": "xxxx",
      "reviewed_name": "xxxx",
      "daily_report": [
        {
          "id": "xxxx",
          "weekly_report_id": "xxxx",
          "uuid_akun": "xxxx",
          "id_reg_penawaran": "xxxx",
          "report_date": "xxx",
          "status": "xxxx",
          "feeling_level": "xxxx",
          "report": "xxxx",
          "string_day": "xxx",
          "counter": "xxxx",
          "created_by": "xxxx",
          "updated_by": "xxxx",
          "created_at": "xxx",
          "updated_at": "xxx"
        }, ...
      ],
      "mentor_notes": [],
      "updated_at": "xxxx"
    }
  ],
  "meta": {
    "limit": "xxxx",
    "offset": "xxxx",
    "total": "xxxx"
  }
}
```
## Tech Stack

**System:** NodeJS, Javascripts, SheetJS




