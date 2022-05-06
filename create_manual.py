
'''

 ______     __  __     ______     __     ______    
/\  ___\   /\ \_\ \   /\  == \   /\ \   /\  ___\   
\ \ \____  \ \  __ \  \ \  __<   \ \ \  \ \___  \  
 \ \_____\  \ \_\ \_\  \ \_\ \_\  \ \_\  \/\_____\ 
  \/_____/   \/_/\/_/   \/_/ /_/   \/_/   \/_____/ 
                                                   

Create Twitch Manual on Google Drive
April 2022

This script takes advantage of the Google Drive API and Google CLoud API
To post up a new manual in form of a Google Slides
Presentation  on Google Drive based on a given set of Fonts,
Videos, and Images.


'''

import os
import datetime
from google.cloud import storage
from __future__ import print_function
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow



# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/presentations',
          'https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/drive.file',
          'https://www.googleapis.com/auth/devstorage.write_only']


def main():
    """Shows basic usage of the Slides API.
    Prints the number of slides and elments in a sample presentation.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    SLIDES = build('slides', 'v1', credentials=creds)
    DRIVE = build('drive', 'v3', credentials=creds)
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "client_secret.json"
    emu = 9525

    # Pass in a query string and a match string to get the item id
    def get_id(query, match):
        while True:

            drive_response = DRIVE.files().list(q=query,
                                                spaces='drive', fields='files(id, name)').execute()

            for file in drive_response.get('files', []):
                if (file.get('name') == match):
                    item_id = file.get('id')
                    return item_id

    # Get input from the user and return it as the main guide name
    def get_guide_name():

        guide_name = input("Enter in the name of the guide: ")
        print(f"The Guide Name is: {guide_name}")
        return guide_name

    # Create a folder in another folder and get the ID
    def create_folder_in_folder(parent_folder_id, folder_name):

        folder_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_folder_id]
        }

        # Create Folder
        folder = DRIVE.files().create(body=folder_metadata, fields='id').execute()
        folder_id = folder.get('id')
        print(f"The Parent Folder ID is: {folder_id}")
        return folder_id

    # Create a presentation in a folder with the given name and return the ID
    def create_guide_in_folder(parent_folder_id, guide_name):
        presentation_metadata = {
            'name': guide_name,
            'mimeType': 'application/vnd.google-apps.presentation',
            'parents': [parent_folder_id]
        }

        # Create the presentation
        DRIVE.files().create(body=presentation_metadata, fields='id').execute()

        presentation_id = get_id("mimeType='application/vnd.google-apps.presentation'", guide_name)
        print(f"The Presentation ID is: {presentation_id}")
        return presentation_id

    # Delete the first slide to start fresh
    def delete_first_slide(guide_id):

        if (guide_id):
            slide_1_id = 'p'
        else:
            print("Presentation is not there")

        # Iterate through the first slide object and look for all the page element 'objectId's
        # We gather this to create a delete requests to delete these items as we don't need these

        deleteObject = [
            {
                "deleteObject": {
                    "objectId": slide_1_id
                }
            }
        ]

        requests = {
            "requests": deleteObject
        }

        # send the delete requests
        SLIDES.presentations() \
            .batchUpdate(presentationId=guide_id, body=requests).execute()
        print("The First Slide Was Deleted")

    # create a slide and get back the ID
    def add_slide(guide_id):
        requests = {
            "requests": [
                {
                    "createSlide": {
                        "slideLayoutReference": {
                            "predefinedLayout": "BLANK"
                        }
                    }
                }
            ]
        }

        slide = SLIDES.presentations() \
            .batchUpdate(presentationId=guide_id, body=requests).execute().get('replies')[0].get('createSlide')
        slide_id = slide.get('objectId')
        return slide_id

    # Create a list Slide ID's and append a Slide ID for every time a new slide is created.
    slide_list = []
    # Create an array of slides based on images provided from a folder. Each image will be uploaded to Google Cloud
    # and a Signed Url will be parsed into the Google Slides API to add the image to each slide.
    # After the image is added, the image will be deleted from Google Cloud

    def create_slides(guide_id):

        bucket_name = "mbpimages"
        storage_client = storage.Client()
        bucket = storage_client.bucket(bucket_name)

        directory = input("Please input the path for the images for this Guide: ")
        stripped_directory = directory.replace('"', "")

        manual_image_path = os.path.normpath(stripped_directory)

        if os.path.isdir(manual_image_path):
            pass
        else:
            print("You did not enter a proper Path. Exiting...")

        for filename in sorted(os.listdir(manual_image_path)):

            slide_id = add_slide(guide_id)
            slide_list.append(slide_id)

            filepath = os.path.join(manual_image_path, filename)
            blob = bucket.blob(filename)

            blob.upload_from_filename(filepath)
            print(f"The following file was uploaded: {filename}")

            image_url = blob.generate_signed_url(
                version="v4",
                expiration=datetime.timedelta(minutes=1),
                method="GET",
                )

            update_requests = [
                {
                    "updatePageProperties": {
                        "objectId": slide_id,
                            "pageProperties": {
                                "pageBackgroundFill": {
                                    "stretchedPictureFill": {
                                        "contentUrl": image_url
                                    }
                                }
                            },
                        "fields": "pageBackgroundFill"
                    }
                }
            ]

            update_slide_body = {
                'requests': update_requests
            }

            # Update the slide by changing the background
            SLIDES.presentations() \
                .batchUpdate(presentationId=guide_id, body=update_slide_body).execute()

            # Delete the uploaded image
            blob.delete()
            print(f"The following file was deleted: {filename}")

    # Add Text to a slide, with size and position paramters based on the presentation and slide ID
    # If a link is provided then the text will be converted to a hypertext link.
    def add_text(presentation_id, slide_id, short_name, text, link, font_size, size_x, size_y, pos_x, pos_y):

        width = {
            'magnitude': size_x * emu,
            'unit': 'EMU'
        }

        height = {
            'magnitude': size_y * emu,
            'unit': 'EMU'
        }

        requests = [
            {
                'createShape': {
                    'objectId': short_name,
                    'shapeType': 'TEXT_BOX',
                    'elementProperties': {
                        'pageObjectId': slide_id,
                        'size': {
                            'width': width,
                            'height': height
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': pos_x * emu,
                            'translateY': pos_y * emu,
                            'unit': 'EMU'
                        }
                    }
                }
            },

            # Insert text into the box, using the supplied element ID.
            {
                'insertText': {
                    'objectId': short_name,
                    'text': text
                }
            },

            {
                'updateTextStyle': {
                    'objectId': short_name,
                    'textRange': {
                        'type': 'ALL',
                    },
                    'style': {
                        'fontFamily': 'Nunito',
                        'fontSize': {
                            'magnitude': font_size,
                            'unit': 'PT'
                        }
                    },
                    'fields': 'fontFamily, fontSize'
                }
            }

        ]

        if (link):
            add_link = {
                'updateTextStyle': {
                        'objectId': short_name,
                        'textRange': {
                            'type': 'ALL',
                        },
                        'style': {
                            'link': {
                                'url': link
                            }
                        },
                        'fields': 'link'
                    }
                }
            requests.append(add_link)
        elif (link is False):
            print("No Link was provided")

        # Execute the request.
        body = {
            'requests': requests
        }

        response = SLIDES.presentations() \
            .batchUpdate(presentationId=presentation_id, body=body).execute()
        create_shape_response = response.get('replies')[0].get('createShape')
        shape_id = create_shape_response.get('objectId')
        print(f"The following text was added with the ID: {shape_id}: {text}")

    # Add a video to a given slide based on the Video ID.
    # The video ID is the ID parameter in a Youtube URL.
    # With the video id we can position the video on the slide
    def add_video(presentation_id, slide_id, short_name, video_id, size_x, size_y, pos_x, pos_y):

        width = {
            'magnitude': size_x * emu,
            'unit': 'EMU'
        }

        height = {
            'magnitude': size_y * emu,
            'unit': 'EMU'
        }

        requests = [
            {
                "createVideo": {
                    "objectId": short_name,
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {
                            "width": width,
                            "height": height
                        },
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": pos_x * emu,
                            "translateY": pos_y * emu,
                            "unit": "EMU"
                        }
                    },
                    "source": "YOUTUBE",
                    "id": video_id
                }
            },

            {
                "updateVideoProperties": {
                    "objectId": short_name,
                    "fields": "autoPlay",
                    "videoProperties": {
                        "autoPlay": True
                    }
                }
            }

        ]

        # Execute the request.
        body = {
            'requests': requests
        }

        response = SLIDES.presentations() \
            .batchUpdate(presentationId=presentation_id, body=body).execute()
        create_video_response = response.get('replies')[0].get('createVideo')
        video_id = create_video_response.get('objectId')
        print(f"The following video ID was added: {short_name}: {video_id}")

    stream_guide_folder_id = get_id("name='Stream_Guides'", "Stream_Guides")

    guide_name = get_guide_name()

    guide_folder = create_folder_in_folder(stream_guide_folder_id, guide_name)

    guide_id = create_guide_in_folder(guide_folder, guide_name)

    # Delete the first generic Slide that gets created automatically
    delete_first_slide(guide_id)

    # Create all the Slides
    create_slides(guide_id)

    # Now go through all the slides and add all the necessary text, videos, and links
    add_text(presentation_id=guide_id,
             slide_id=slide_list[3],
             short_name="font_description",
             text="Download and Install all the following fonts to your computer:",
             link=False,
             font_size=18,
             size_x=720,
             size_y=50,
             pos_x=120,
             pos_y=130)

    # Get input from the user for all the available Fonts for this package
    count = 1
    y_position = 180
    while True:
        font_name = input("Enter the name of the font: ")
        font_url = input(f"Enter in the URL for the font: {font_name}: ")

        add_text(presentation_id=guide_id,
                 slide_id=slide_list[3],
                 short_name=f"font_{str(count)}",
                 text=f"{font_name}",
                 link=f"{font_url}",
                 font_size=16,
                 size_x=720,
                 size_y=50,
                 pos_x=120,
                 pos_y=y_position)

        count += 1
        y_position += 40
        ask_user = input("Do you have any more fonts to add?[yes/no]: ")
        if ask_user == "yes" or ask_user == "y":
            continue
        else:
            break

    # Add the How To OBS Install Video
    add_video(presentation_id=guide_id,
              slide_id=slide_list[4],
              short_name="obs_install",
              video_id="Gr4XgEt2eXM",
              size_x=600,
              size_y=337,
              pos_x=180,
              pos_y=110)

    # Add the How To SLOBS Install Video
    add_video(presentation_id=guide_id,
              slide_id=slide_list[5],
              short_name="slobs_install",
              video_id="5X--doqvyRE",
              size_x=600,
              size_y=337,
              pos_x=180,
              pos_y=110)

    # Add the How To Stinger to OBS Video
    add_video(presentation_id=guide_id,
              slide_id=slide_list[6],
              short_name="stinger_obs",
              video_id="7KtAZTZ8rtM",
              size_x=600,
              size_y=337,
              pos_x=180,
              pos_y=110)

    get_trans_point = input("Input a number for the Frame Transition Point: ")

    if get_trans_point == int:
        frame_point = f"Frame Transition Point = {get_trans_point}"
    else:
        get_new_point = input(f"You did not input a number. Did you want to add: {get_trans_point}?[yes/no]")
        if get_new_point == "yes" or get_new_point == "y":
            frame_point = f"{get_trans_point}"

    add_text(presentation_id=guide_id,
            slide_id=slide_list[6],
            short_name="trans_obs",
            text=f"{frame_point}",
            link=False,
            font_size=16,
            size_x=720,
            size_y=50,
            pos_x=360,
            pos_y=450)

    add_video(presentation_id=guide_id,
            slide_id=slide_list[7],
            short_name="stinger_slobs",
            video_id="cEMJAiRE47o",
            size_x=600,
            size_y=337,
            pos_x=180,
            pos_y=110)

    add_text(presentation_id=guide_id,
            slide_id=slide_list[7],
            short_name="trans_slobs",
            text=f"{frame_point}",
            link=False,
            font_size=16,
            size_x=720,
            size_y=50,
            pos_x=360,
            pos_y=450)

    add_text(presentation_id=guide_id,
            slide_id=slide_list[8],
            short_name="alert_description_1",
            text="Please follow instructions from StreamLabs to setup your Alerts!",
            link=False,
            font_size=18,
            size_x=720,
            size_y=50,
            pos_x=120,
            pos_y=110)

    add_text(presentation_id=guide_id,
            slide_id=slide_list[8],
            short_name="streamlabs_link",
            text="Setting up your Alerts",
            link="https://streamlabs.com/content-hub/post/setting-up-your-streamlabs-alerts",
            font_size=18,
            size_x=720,
            size_y=50,
            pos_x=120,
            pos_y=180)

    add_text(presentation_id=guide_id,
            slide_id=slide_list[8],
            short_name="alert_description_2",
            text="Add your alerts into StreamLabs online and you're up and ready!",
            link=False,
            font_size=18,
            size_x=720,
            size_y=50,
            pos_x=120,
            pos_y=240)

    add_text(presentation_id=guide_id,
            slide_id=slide_list[8],
            short_name="alert_description_3",
            text="Feel free to customize them as needed!",
            link=False,
            font_size=18,
            size_x=720,
            size_y=50,
            pos_x=120,
            pos_y=270)


if __name__ == '__main__':
    main()