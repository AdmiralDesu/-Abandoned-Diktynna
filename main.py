import imapclient
import smtplib
import pandas as pd
import copy
import os.path
import PySimpleGUI as sg
from email.message import EmailMessage
import email
from xlsxwriter.utility import xl_col_to_name


class Andromeda:

    def __init__(self):
        """
        unparsed_emails: почты, на которые нужно отправить письма
        accounts: почтовые аккаунты, с которых отправляются письма
        text: шаблоны писем
        parsed_emails: почты, на которые письма были отправлены
        """
        print('Андромеда начала работу.')
        self.unparsed_emails = list()
        self.accounts = dict()
        self.parsed_emails = dict()
        self.message_template = EmailMessage()
        self.message = EmailMessage()
        # Сверху создаются контейнеры для будущей работы
        # Сверху создаются IMAP(для изучения почты) и SMTP(для отправки сообщений) сервера

    def emails_handler(self, path_to_emails_excel='./input/emails.xlsx'):

        print('Начинаю обработку почт.')
        xl_file = pd.read_excel(path_to_emails_excel, na_filter=False)
        xl_file.fillna(0)
        emails = list(filter(lambda item: item != '', xl_file['Почты'].tolist()))

        counties = list(filter(lambda item: item != '', xl_file['Страна'].tolist()))
        for index, item in enumerate(emails):
            if counties[index] not in ('US', 'GB', 'CA', 'RUS', 'BEL', 'UKR'):
                self.unparsed_emails.append(item)

        print(f'Почты схвачены:\n{self.unparsed_emails}')

    def accounts_handler(self, path_to_accounts_excel='./input/accounts.xlsx'):

        print('Начинаю обработку аккаунтов.')
        xl_file = pd.read_excel(path_to_accounts_excel, na_filter=False)

        self.accounts = dict.fromkeys(list(filter(lambda item: item != '', xl_file['Почты'].tolist())))
        passwords = list(filter(lambda item: item != '', xl_file['Пароли'].tolist()))
        names = list(filter(lambda item: item != '', xl_file['Имя'].tolist()))

        for index, key in enumerate(self.accounts.keys()):
            self.accounts[key] = [passwords[index], names[index]]
        self.parsed_emails = dict.fromkeys(list(filter(lambda item: item != '', xl_file['Почты'].tolist())))
        for key in self.parsed_emails.keys():
            self.parsed_emails[key] = list()
        print(f'Аккаунты схвачены:\n{self.accounts}')

    def text_handler(self, subject, text):
        print('Создаю шаблон...')
        del self.message_template['Subject']
        self.message_template['Subject'] = subject
        self.message_template.set_content(text)
        print(f'Шаблон создан\n{self.message_template}')

    def form_message(self, sender, name=None):
        self.message = copy.deepcopy(self.message_template)
        print(sender, name)
        if '*name*' in self.message.get_content():
            contest_list = self.message.get_content().split(' ')
            for item in contest_list:
                if item == '*name*,':
                    contest_list[contest_list.index(item)] = f'{name},'
                if item == '*name*!':
                    contest_list[contest_list.index(item)] = f'{name}!'
                if item == '*name*.':
                    contest_list[contest_list.index(item)] = f'{name}.'
                if item == '*name*':
                    contest_list[contest_list.index(item)] = f'{name}'

            named_content = ' '.join(contest_list)
            self.message.set_content(named_content)
            print(fr'{self.message.get_content()}')

        self.message['From'] = sender
        self.message['To'] = self.unparsed_emails[0]
        self.message['Date'] = email.utils.formatdate()
        self.message['Message-ID'] = email.utils.make_msgid(domain='uniqlomanagement.com')

    def save_result(self):
        writer = pd.ExcelWriter('./Output/parsed_emails.xlsx', engine='xlsxwriter')
        parsed_emails_to_save = pd.DataFrame.from_dict(self.parsed_emails)
        parsed_emails_to_save.to_excel(writer, sheet_name='Sheet1')
        worksheet = writer.sheets['Sheet1']
        worksheet.set_column(f'{xl_col_to_name(1)}:{xl_col_to_name(len(self.parsed_emails.keys()))}', 30)

        writer.save()

    def mailman(self, limit=7):
        """

        :param limit: лимит сообщений с одной почты
        :return: None
        Вводится счетчик сообщений messages
        Если отправки уже совершались, то обработанные почты считываются с excel файла
        Далее идет проверка, были ли внесенные почты уже обработаны, если да, то они не рассматриваются
        Идет отправка сообщений с каждой почты
        Выводится количество отправленных сообщений
        В отдельный файл excel выводятся обработанные почты

        """

        out_of_emails = False
        print('Почтальон начал работу')
        messages = 0
        if os.path.exists('./Output/parsed_emails.xlsx'):
            xl_file = pd.read_excel('./Output/parsed_emails.xlsx')
            for key in self.parsed_emails.keys():
                self.parsed_emails[key] = xl_file[key].tolist()
            if 'Unnamed: 0' in self.parsed_emails.keys():
                self.parsed_emails.pop('Unnamed: 0')
            else:
                pass

        for key in self.parsed_emails.keys():
            self.unparsed_emails = list(filter(lambda item: item not in self.parsed_emails[key], self.unparsed_emails))
            self.parsed_emails[key] = list(filter(lambda item: item == 'Не хватило почты', self.unparsed_emails))
        print('Начинаю рассылку...')
        if len(self.unparsed_emails) == 0:
            print('Почт нет')
        else:
            for sender in self.accounts.keys():
                if out_of_emails is True:
                    break
                i = 1
                while i <= limit or out_of_emails is True:

                    server_smtp = smtplib.SMTP_SSL('smtp.mail.ru', 465)
                    server_smtp.login(sender, self.accounts[sender][0])

                    self.form_message(sender, self.accounts[sender][1])

                    print(f'Сообщение от {sender} отправляется {self.unparsed_emails[0]}.')

                    server_smtp.send_message(self.message)

                    print('Сообщение отправлено.')
                    self.parsed_emails[sender].append(self.unparsed_emails[0])
                    self.unparsed_emails.pop(0)

                    print(f'Оставшиеся почты\n{self.unparsed_emails}')
                    print(f'Обработанные почты\n{self.parsed_emails}')
                    messages += 1
                    i += 1
                    server_smtp.quit()

                    server_imap = imapclient.IMAPClient('imap.mail.ru', ssl=True)
                    server_imap.login(sender, self.accounts[sender][0])

                    server_imap.append('Отправленные', self.message.as_string())
                    if len(self.unparsed_emails) == 0:
                        print('Почты закончились')
                        out_of_emails = True
                        break

        print(f'Почтальон закончил работу\nБыло отправлено {messages} сообщений')
        print(self.parsed_emails)
        for key in self.parsed_emails.keys():
            while len(self.parsed_emails[key]) < len(max(self.parsed_emails.values())):
                self.parsed_emails[key].append('Не хватило почты')
        print(self.parsed_emails)
        self.save_result()
        print('Данные сохранены')

    def interface_and_work(self):

        emails_part = sg.Frame('To', [[sg.Text('Path to emails excel', justification='c')],
                                      [sg.Input(size=(20, 1), key='Path_to_emails_excel'), sg.FileBrowse()],
                                      [sg.Button('Parse', key='Parse_emails')]])

        accounts_part = sg.Frame('From', [[sg.Text('Path to accounts excel', justification='c')],
                                          [sg.Input(size=(20, 1), key='Path_to_accounts_excel'),
                                           sg.FileBrowse()],
                                          [sg.Button('Parse', key='Parse_accounts')]])

        layout = [
            [sg.Text('Andromeda v0.8', justification='c', font='TimesNewRoman 16')],
            [sg.HSeparator()],
            [emails_part, accounts_part],
            [sg.HSeparator()],
            [sg.Frame('Letter', [[sg.Text('Head of the letter', justification='c')],
                                 [sg.Input(size=(20, 1), key='Head_of_the_letter'),
                                  sg.Text('Limit of letters form 1 email'),
                                  sg.Input(key='Limit', size=(3, 1))],
                                 [sg.Text('Text of the letter', justification='c')],
                                 [sg.Multiline(size=(60, 20), key='Text_of_the_letter')],
                                 [sg.Button('Use this template', key='Text')]
                                 ])
             ],
            [sg.Button('Start mailing', key='Start'), sg.Button('Exit', key='Exit')],
            [sg.Frame('Console', [[sg.Multiline('Андромеда начала работу\nПодключения к IMAP и SMTP '
                                                'установлены', size=(60, 8), key='Console')]])]
        ]

        window = sg.Window('Andromeda', layout, element_justification='c', resizable=True)
        event_chain = 'Андромеда начала работу\nПодключения к IMAP и SMTP установлены'
        while True:
            event, values = window.read(timeout=100)
            # print(event, values)

            if event == 'Parse_emails':
                event_chain += '\nНачинаю обработку почт'
                window['Console'].update(event_chain)
                self.emails_handler(values['Path_to_emails_excel'])
                event_chain += '\nПочты схвачены'
                window['Console'].update(event_chain)
            elif event == 'Parse_accounts':
                event_chain += '\nНачинаю обработку аккаунтов'
                window['Console'].update(event_chain)
                self.accounts_handler(values['Path_to_accounts_excel'])
                event_chain += '\nАккаунты схвачены'
                window['Console'].update(event_chain)
            elif event == 'Text':
                event_chain += '\nСоздается шаблон'
                window['Console'].update(event_chain)
                self.text_handler(values['Head_of_the_letter'], values['Text_of_the_letter'])

                event_chain += '\nШаблон создан'
                window['Console'].update(event_chain)
            elif event == 'Start':
                event_chain += '\nНачинаю рассылку писем'
                window['Console'].update(event_chain)
                print(f"{int(values['Limit'])} - это лимит.")
                self.mailman(int(values['Limit']))
                event_chain += '\nРассылка писем окончена'
                window['Console'].update(event_chain)
            if event in (None, 'Exit'):
                window.close()
                break


rocket = Andromeda()
rocket.interface_and_work()

