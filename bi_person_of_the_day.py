from secrets import secrets as sec
import pandas as pd
import sys
import os
import slack


print('---------------------------')
print('Start: BI Person of the Day')
print('---------------------------')


def get_person():
    df = pd.read_excel('bi_person_of_the_day.xlsx')
    df = df.loc[df['available'] == 'Y']
    df = df.sample()
    bi_person_name = df['name'].values[0]
    bi_person_id = df['slack_id'].values[0]
    print(f"Success: BI Person of the Day chosen: {bi_person_name}")

    return bi_person_name, bi_person_id


def send_slack(bi_person_name, bi_person_id, channel='#ask_bi'):
    os.environ['SLACK_API_TOKEN'] = sec.slack_token
    client = slack.WebClient(token=os.environ['SLACK_API_TOKEN'])
    print("Success: Slack Authentication successful")

    response = client.chat_postMessage(
        channel=f'{channel}',
        link_names=1,
        text=f'''<@{bi_person_id}> {bi_person_name} is our BI Person of the Day!'''
        )

    return response


if __name__ == '__main__':
    try:
        bi_person_name, bi_person_id = get_person()
        send_slack(bi_person_name, bi_person_id, channel='#test_channel')
        print(f"Success: Slack Message sent to {bi_person_name}")

    except Exception as e:
        print(f"Failure: {e}")
        sys.exit()

    finally:
        print('----------------------------')
        print('Finish: BI Person of the Day')
        print('----------------------------')
        sys.exit()
