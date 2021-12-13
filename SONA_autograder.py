#!/usr/bin/env python
# coding: utf-8

import email
import imaplib
import os
import time

from dotenv import load_dotenv
from loguru import logger
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


def main(driver_path='/usr/bin/chromedriver'):
    load_dotenv()
    logger.add('SONA.log')

    service = Service(driver_path)
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('headless')
    driver = webdriver.Chrome(options=options, service=service)

    driver.get('https://unomaha.sona-systems.com/Default.aspx?ReturnUrl=/')

    login = driver.find_element(By.ID,
                                'ctl00_ContentPlaceHolder1_pnlStandardLogin')

    for e in login.find_elements(By.TAG_NAME, 'input'):
        if e.get_attribute('name') == 'ctl00$ContentPlaceHolder1$userid':
            e.send_keys(os.environ['USERNAME'])
        elif e.get_attribute('name') == 'ctl00$ContentPlaceHolder1$pw':
            e.send_keys(os.environ['PASSWD'])
            time.sleep(0.5)
            e.send_keys(Keys.RETURN)
            break

    driver.get('https://unomaha.sona-systems.com/uncredited_slots.aspx')
    time.sleep(1)

    participants = []

    for row in driver.find_elements(By.TAG_NAME, 'tr')[:-1]:
        for cell in row.find_elements(By.TAG_NAME, 'td'):
            if cell.get_attribute('data-title') == 'Participant':
                name = cell.find_element(By.TAG_NAME, 'span').text
            elif cell.get_attribute('data-title') == 'Action':
                for action in cell.find_elements(By.TAG_NAME, 'input'):
                    if action.get_attribute('value') == 'GrantRadioButton':
                        participants.append((name, action))
                        break

    logger.info(f'Found {len(participants)} uncredited timeslots.')

    if len(participants) == 0:
        logger.info('No uncredited timeslots. Exiting...')
        driver.quit()
        return

    mail = imaplib.IMAP4_SSL('outlook.office365.com')
    mail.login(os.environ['EMAIL_ADDR'], os.environ['EMAIL_PASSWD'])
    mail.select('INBOX')
    _, selected_mails = mail.search(None, '(FROM "noreply@qemailserver.com")')
    email_ids = selected_mails[0].split()

    for email_id in email_ids[(len(participants) + 100) * -1:]:
        _, data = mail.fetch(email_id, '(RFC822)')
        _, bytes_data = data[0]
        email_message = email.message_from_bytes(bytes_data)
        granted_to = []
        for part in email_message.walk():
            if part.get_content_type(
            ) == 'text/plain' or part.get_content_type() == 'text/html':
                body = part.get_payload(decode=True).decode()
                for participant in participants:
                    if participant[0].lower() in body.lower():
                        participant[1].click()
                        logger.info(f'Granting credit to {participant[0]}...')
                        granted_to.append(participant)
    time.sleep(1)

    for e in driver.find_elements(By.TAG_NAME, 'input'):
        if e.get_attribute('type') == 'submit':
            e.click()
            logger.info(
                f'Finished! Granted credits to {len(granted_to)} participant(s).'
            )
            break

    driver.quit()


if __name__ == '__main__':
    main()
