import requests

api_url_base = 'http://itunes.apple.com/lookup'
headers = {'Content-Type': 'application/json'}


def get_app_data_by_id(id):
    id = str(id)
    response = requests.get(api_url_base + "?id=" + id)

    if response.status_code == 200:
        return response.json()
    else:
        print("request get failed for id: {} | error: {}".format(id, response.status_code))
