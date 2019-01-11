#!/usr/bin/env python3
"""
Export all tasks from my Asana workspaces to Google Spreadsheet
Requires asana library.  See https://github.com/Asana/python-asana
"""
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

import asana

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/drive.metadata.readonly https://www.googleapis.com/auth/spreadsheets.readonly'

# replace with your personal asana access token.
INVITEE_PERSONAL_ACCESS_TOKEN = '0/2d89202653fc4a83a075e2136f09397b'
OWNER_PERSONAL_ACCESS_TOKEN = '0/894b80abb9f5a42d6f257c3522750db8'

PERSONAL_ACCESS_TOKEN = OWNER_PERSONAL_ACCESS_TOKEN

SPREADSHEET_FOLDER = "tasks"
SPREADSHEET_NAME_PREFIX = "Tasks"

HEADERS = ['Workspace', 'Workspace Id', 'Project', 'Project Id', 'Sync', 'Task', 'Task Id', 'Parent', 'Due Date', 'Created At', \
    'Modified At', 'Completed', 'Completed At', 'Assignee', 'Assignee Status', \
    'Notes']

def process_project_tasks(client, project):
    """Add each task for the current project to the records list."""

    headers = {'Authorization': 'Bearer ' + PERSONAL_ACCESS_TOKEN}
    import requests

    payload = {'resource': project['gid']}
    response = requests.get('https://app.asana.com/api/1.0/events', headers=headers, params=payload)
    response_json = response.json()
    if response_json['errors']:
      payload = {'resource': project['gid'], 'sync': response_json['sync']}
      response = requests.get('https://app.asana.com/api/1.0/events', headers=headers, params=payload)
      response_json = response.json()

    task_list = []
    while True:
        tasks = client.tasks.find_by_project(project['id'], {"opt_fields":"name, \
            projects, workspace, id, due_on, created_at, modified_at, completed, \
            completed_at, assignee, assignee_status, parent, notes"})
        for task in tasks:
            assignee = task['assignee']['id'] if task['assignee'] is not None else ''
            created_at = task['created_at'][0:10] + ' ' + task['created_at'][11:16] if \
                    task['created_at'] is not None else None
            modified_at = task['modified_at'][0:10] + ' ' + task['modified_at'][11:16] if \
                    task['modified_at'] is not None else None
            completed_at = task['completed_at'][0:10] + ' ' + task['completed_at'][11:16] if \
                task['completed_at'] is not None else None
            rec = [project['name'] if len(task_list) == 0 else None
              , project['gid'] if len(task_list) == 0 else None
              , response_json['sync'] if len(task_list) == 0 else None
              , task['name'], task['id'], task['parent']
              , task['due_on'], created_at,
              modified_at, task['completed'], completed_at, assignee,
              task['assignee_status'], task['notes']
            ]
            rec = ['' if s is None else s for s in rec]
            task_list.append(rec)

        if 'next_page' not in tasks:
            break
    return task_list

def main():
  try:
    # Construct an Asana client
    client_asana = asana.Client.access_token(PERSONAL_ACCESS_TOKEN)

    client_asana.options['page_size'] = 100

    # Get your user info
    me = client_asana.users.me()

  except asana.error.NoAuthorizationError as err:
    print(err)
  else:
    tasks = []
    # For each workspace, iterate through all projects and tasks
    for workspace in me['workspaces']:
      all_projects = client_asana.projects.find_by_workspace(workspace['gid'], iterator_type=None)
      tasks_in_workspace = []
      for project in all_projects:
        tasks_in_workspace.extend(process_project_tasks(client_asana, project))
      for task_index, task in enumerate(tasks_in_workspace):
        if task_index == 0:
          task.insert(0, workspace['gid'])
          task.insert(0, workspace['name'])
        else:
          task.insert(0, "")
          task.insert(0, "")
      tasks.extend(tasks_in_workspace)

    # identify workspace column, and other columns to hide
    workspace_column = None
    workspace_id_column = None
    project_id_column = None
    task_id_column = None
    sync_column = None
    import re
    for index,header in enumerate(HEADERS):
      if re.search("(^|\W)project(\W|$)", header, re.I) and re.search("(^|\W)id(\W|$)", header, re.I):
        project_id_column = index
      elif re.search("(^|\W)task(\W|$)", header, re.I) and re.search("(^|\W)id(\W|$)", header, re.I):
        task_id_column = index
      elif re.search("sync", header, re.I):
        sync_column = index
      elif re.search("(^|\W)workspace(\W|$)", header, re.I) and re.search("(^|\W)id(\W|$)", header, re.I):
        workspace_id_column = index
      elif re.search("workspace", header, re.I):
        workspace_column = index

    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    store = file.Storage('token.json')
    creds = store.get()
    flags = tools.argparser.parse_args(args=[])
    if not creds or creds.invalid:
      flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
      creds = tools.run_flow(flow, store, flags)

    service = build('drive', 'v3', http=creds.authorize(Http()))

    # search for folder (containing sheet)
    searchFolderQuery = u"name = '{}' and mimeType = '{}' and trashed = false".format(SPREADSHEET_FOLDER, "application/vnd.google-apps.folder")
    results = service.files().list(q=searchFolderQuery, pageSize=1, fields="files(id)").execute()
    sheetFolders = results.get('files', [])

    # folder does not exist
    if not sheetFolders:
      folder_metadata = {
          'name': SPREADSHEET_FOLDER,
          'mimeType': 'application/vnd.google-apps.folder'
      }
      folder = service.files().create(body=folder_metadata,
                                          fields='id').execute()
      sheetFolderId = folder.get("id")
    else:
      sheetFolderId = sheetFolders[0].get("id")

    spread_sheet_name = SPREADSHEET_NAME_PREFIX + " - " + me["name"]
    sheet_name = spread_sheet_name

    # search sheet in folder
    searchSheetQuery = u"name = '{}' and mimeType = '{}' and '{}' in parents and trashed = false".format(spread_sheet_name, "application/vnd.google-apps.spreadsheet", sheetFolderId)
    results = service.files().list(q=searchSheetQuery, pageSize=1, fields="files(id)").execute()
    sheets = results.get('files', [])

    # sheet does not exist
    spreadSheetAlreadyExists = False
    if not sheets:
      sheet_metadata = {
          'name': spread_sheet_name,
          'mimeType': 'application/vnd.google-apps.spreadsheet',
          'parents': [sheetFolderId]
      }
      sheet = service.files().create(body=sheet_metadata,
                                          fields='id').execute()
      spreadSheetId = sheet.get("id")
    else:
      spreadSheetAlreadyExists = True
      spreadSheetId = sheets[0].get("id")

    service = build('sheets', 'v4', http=creds.authorize(Http()))

    sheets = service.spreadsheets()

    spreadsheet = sheets.get(spreadsheetId=spreadSheetId, includeGridData = False).execute()

    sheet1Id = -1
    sheetId = -1

    rowGroups = None
    bandedRanges = None

    for sheet in spreadsheet.get("sheets"):
      currentSheetId = sheet.get("properties").get("sheetId")
      currentSheetNameLowerCase = sheet.get("properties").get("title").lower()
      if currentSheetNameLowerCase == "sheet1":
        sheet1Id = currentSheetId
      if currentSheetNameLowerCase == sheet_name.lower():
        sheetId = currentSheetId
        rowGroups = sheet.get("rowGroups")
        bandedRanges = sheet.get("bandedRanges")

    # sheet does not exist
    if sheetId == -1:
      batch_update_spreadsheet_request_body = {
        "includeSpreadsheetInResponse" : False
        , 'requests' : [{
            "addSheet" : {
              "properties" : {
                "title" : sheet_name
                , "tabColor": {
                  "red": 0.3,
                  "green": 1.0,
                  "blue": 0.6
                }
              }
            }
          }
        ]
      }
      response = sheets.batchUpdate(spreadsheetId=spreadSheetId, body=batch_update_spreadsheet_request_body).execute()
      sheetId = response.get("replies")[0].get("addSheet").get("properties").get("sheetId")

    # remove default sheet, if spreadsheet is new
    if not spreadSheetAlreadyExists and sheet1Id != -1:
      batch_update_spreadsheet_request_body = {
        "includeSpreadsheetInResponse" : False
        , 'requests' : [{
            "deleteSheet" : {
              "sheetId" : sheet1Id
            }
          }
        ]
      }

      sheets.batchUpdate(spreadsheetId=spreadSheetId, body=batch_update_spreadsheet_request_body).execute()

    # header + data + an extra
    sheetRowsCount = len(tasks) + 2

    # multiple requests: 
    # set count of rows and columns
    # freeze header
    # hide workspace id column
    # hide project id column
    # hide task id column
    # hide sync column
    # clear sheet
    # add header and tasks
    # format header
    # add border to header
    batch_update_spreadsheet_request_body = {
      "includeSpreadsheetInResponse" : False
      , 'requests' : [
        {
          "updateSheetProperties" : {
            "properties" : {
              "sheetId": sheetId,
              "gridProperties" : {
                "rowCount" : sheetRowsCount
                , "columnCount" : len(HEADERS)
              }
            },
            "fields" : "gridProperties.rowCount, gridProperties.columnCount"
          }
        }
        , {
          "updateSheetProperties" : {
            "properties" : {
              "sheetId": sheetId,
              "gridProperties" : {
                "frozenRowCount" : 1
              }
            }
            , "fields" : "gridProperties.frozenRowCount"
          }
        }
        , {
          "updateDimensionProperties" : {
            "range" : {
              "sheetId" : sheetId
              , "dimension" : "COLUMNS"
              , "startIndex" : workspace_id_column
              , "endIndex" : workspace_id_column + 1
            }
            , "properties" : {
              "hiddenByUser": True
            }
            , "fields" : "hiddenByUser"
          }
        }
        , {
          "updateDimensionProperties" : {
            "range" : {
              "sheetId" : sheetId
              , "dimension" : "COLUMNS"
              , "startIndex" : project_id_column
              , "endIndex" : project_id_column + 1
            }
            , "properties" : {
              "hiddenByUser": True
            }
            , "fields" : "hiddenByUser"
          }
        }
        , {
          "updateDimensionProperties" : {
            "range" : {
              "sheetId" : sheetId
              , "dimension" : "COLUMNS"
              , "startIndex" : task_id_column
              , "endIndex" : task_id_column + 1
            }
            , "properties" : {
              "hiddenByUser": True
            }
            , "fields" : "hiddenByUser"
          }
        }
        , {
          "updateDimensionProperties" : {
            "range" : {
              "sheetId" : sheetId
              , "dimension" : "COLUMNS"
              , "startIndex" : sync_column
              , "endIndex" : sync_column + 1
            }
            , "properties" : {
              "hiddenByUser": True
            }
            , "fields" : "hiddenByUser"
          }
        }
        , {
          "updateCells" : {
            "rows" : []
            , "fields" : "userEnteredValue"
            , "start" : {
              "sheetId": sheetId,
              "rowIndex": 0,
              "columnIndex": 0
            }
          }
        }
        , {
          "updateCells" : {
            "rows" : []
            , "fields" : "userEnteredValue"
            , "start" : {
              "sheetId": sheetId,
              "rowIndex": 0,
              "columnIndex": 0
            }
          }
        }
        , {
          "repeatCell": {
            "range": {
              "sheetId": sheetId,
              "startRowIndex": 0,
              "endRowIndex": 1
            },
            "cell": {
              "userEnteredFormat": {
                "backgroundColor": {
                  "red": 0.0,
                  "green": 0.0,
                  "blue": 0.0
                },
                "horizontalAlignment" : "CENTER",
                "textFormat": {
                  "foregroundColor": {
                    "red": 1.0,
                    "green": 1.0,
                    "blue": 1.0
                  },
                  "fontSize": 12,
                  "bold": True
                }
              }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
          }
        }
        , {
          "updateBorders": {
            "range": {
              "sheetId": sheetId,
              "startRowIndex": 0,
              "endRowIndex": 1
            },
            "top": {
              "style": "SOLID",
              "color": {
                "blue": 1.0
              },
            },
            "bottom": {
              "style": "SOLID",
              "color": {
                "blue": 1.0
              },
            },
            "innerHorizontal": {
              "style": "SOLID",
              "color": {
                "blue": 1.0
              },
            },
          }
        }
      ]
    }

    if rowGroups:
      for rowGroup in rowGroups:
        # delete groups before setting number of rows and columns
        batch_update_spreadsheet_request_body["requests"].insert(0, {
          "deleteDimensionGroup" : {
            "range" : rowGroup["range"]
          }
        })

    batch_request_clear_cells = None
    batch_request_fill_cells = None
    for request_index, batch_request in enumerate(batch_update_spreadsheet_request_body["requests"]):
      if "updateCells" in batch_request:
        batch_request_clear_cells = request_index
        batch_request_fill_cells = batch_request_clear_cells + 1
        break;

    if batch_request_clear_cells:
      # clear sheet including headers
      for taskRow in range(sheetRowsCount):
        rowCellsBlank = []
        for column in range(len(HEADERS)):
          rowCellsBlank.append({
            "userEnteredValue" : {
              "stringValue" : ""
            }
          })
        batch_update_spreadsheet_request_body["requests"][batch_request_clear_cells]["updateCells"]["rows"].append({"values" : rowCellsBlank})

    if batch_request_fill_cells:
      # headers
      headerCellValues = []
      for header in HEADERS:
        headerCellValues.append({
          "userEnteredValue" : {
            "stringValue" : header
          }
        })

      batch_update_spreadsheet_request_body["requests"][batch_request_fill_cells]["updateCells"]["rows"].append({"values" : headerCellValues})

      workspaceGroup = None
      projectGroup = None

      # tasks
      for task_index, task in enumerate(tasks):
        workspace = task[workspace_column].strip()
        projectId = task[project_id_column ].strip()

        if workspace:
          if workspaceGroup:
            workspaceGroup["addDimensionGroup"]["range"]["endIndex"] = task_index + 1
          workspaceGroup = {
                            "addDimensionGroup" : {
                              "range" : {
                                "sheetId": sheetId
                                , "dimension" : "ROWS"
                                , "startIndex" : task_index + 2
                              }
                            }
                          }
          batch_update_spreadsheet_request_body["requests"].append(workspaceGroup)

        if projectId:
          if projectGroup:
            projectGroup["addDimensionGroup"]["range"]["endIndex"] = task_index + 1
          projectGroup = {
                            "addDimensionGroup" : {
                              "range" : {
                                "sheetId": sheetId
                                , "dimension" : "ROWS"
                                , "startIndex" : task_index + 2
                              }
                            }
                          }
          batch_update_spreadsheet_request_body["requests"].append(projectGroup)

        if task_index == len(tasks) - 1:
          if workspaceGroup:
            workspaceGroup["addDimensionGroup"]["range"]["endIndex"] = task_index + 2
          if projectGroup:
            projectGroup["addDimensionGroup"]["range"]["endIndex"] = task_index + 2

        taskCellsValues = []
        for cellValue in task:
          # boolean
          if type(cellValue) == type(True):
            if cellValue:
              stringValue = "Yes"
            else:
              stringValue = "No"
          else:
            stringValue = str(cellValue)
          taskCellsValues.append({
            "userEnteredValue" : {
              "stringValue" : stringValue
            }
          })
        batch_update_spreadsheet_request_body["requests"][batch_request_fill_cells]["updateCells"]["rows"].append({"values" : taskCellsValues})

    if not bandedRanges:
      batch_update_spreadsheet_request_body["requests"].append({
        "addBanding" : {
          "bandedRange" : {
            "range" : {
              "sheetId" : sheetId
              , "startRowIndex" : 1
            }
            , "rowProperties" : {
              "firstBandColor": {
                "red" : 1.0
                , "green" : 0.89
                , "blue" : 0.88
                , "alpha" : 0.4
              },
              "secondBandColor": {
                "red" : 1.0
                , "green" : 0.97
                , "blue" : 0.86
                , "alpha" : 0.3
              }
            }
          }
        }
      })

    sheets.batchUpdate(spreadsheetId=spreadSheetId, body=batch_update_spreadsheet_request_body).execute()

    print("Sheet id is " + str(spreadSheetId))

if __name__ == '__main__':
    main()