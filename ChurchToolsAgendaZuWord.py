import logging
import tkinter
from datetime import datetime

import docx


class ChurchToolsAgendaZuWord:
    def __init__(self, api):
        self.lbx1 = None
        self.win = None
        self.api = api
        logging.basicConfig(filename='logs/main.py.log', encoding='utf-8',
                            format="%(asctime)s %(name)-10s %(levelname)-8s %(message)s",
                            level=logging.DEBUG)
        logging.debug("ChurchToolsAgendaZuWord initialized")
        self.events = []
        self.event_agendas = []
        self.load_events_with_agenda()
        self.create_gui()

    def load_events_with_agenda(self):
        events_temp = self.api.get_events()
        logging.debug("{} Events loaded".format(len(events_temp)))

        for event in events_temp:
            agenda = self.api.get_event_agenda(event['id'])
            if agenda is not None:
                self.event_agendas.append(agenda)
                self.events.append(event)
        logging.debug("{} Events kept because schedule exists".format(len(events_temp)))

    def create_gui(self):
        win = tkinter.Tk()
        win.title = 'Bitte Event auswählen'

        lbl1 = tkinter.Label(win, text="Die nächsten Veranstaltungen:")
        lbl1.pack()

        lbx1 = tkinter.Listbox(win, width=500)
        i = 0
        for event in self.events:
            startdate = datetime.fromisoformat(event['startDate'][:-1])
            datetext = startdate.__format__('%a %b %d\t%H:%M')
            lbx1.insert(i, datetext + '\t' + event['name'])
            i += 1
        lbx1.pack()

        btn1 = tkinter.Button(win, text='Veranstaltung als Text umwandeln', command=self.btn1_press)
        logging.debug("GUI Elements defined")

        btn1.pack()

        logging.debug("GUI Mainloop started")

        self.win = win
        self.lbx1 = lbx1

        win.mainloop()

    def btn1_press(self):
        logging.debug("Button 1 pressed")
        if len(self.lbx1.curselection()) == 0:
            logging.info("No item selected")
            return
        else:
            event = self.events[self.lbx1.curselection()[0]]
            logging.debug("Selected event ID: {}".format(event['id']))
            #TODO #4 add FileChooser for destination path
            self.process_agenda(self.event_agendas[self.lbx1.curselection()[0]])
            self.win.destroy()

    def process_agenda(self, agenda):
        logging.debug('Processing: ' + agenda['name'])

        document = docx.Document()
        heading = agenda['name']
        heading += '- Entwurf' if not agenda['isFinal'] else ''
        document.add_heading(heading)
        modifiedDate = datetime.strptime(agenda["meta"]['modifiedDate'], '%Y-%m-%dT%H:%M:%S%z')
        modifiedDate2 = modifiedDate.astimezone().strftime('%a %d.%m (%H:%M:%S)')
        document.add_paragraph("Abzug aus CT mit Änderungen bis inkl.: " + modifiedDate2)

        for item in agenda["items"]:
            #todo #3 check if pre_event items should be skipped
            title = str(item["position"]) + ' ' + item["title"]
            if item['type'] == 'song':
                title += ': ' + item['song']['title'] + ' (' + item['song'][
                    'category'] + ')'  # TODO #5 check if fails on empty song items
            document.add_heading(title, level=2)

            resonsible_list = [item['responsible']] if isinstance(item['responsible'], dict) else item['responsible']
            responsible_list = []
            for responsible_item in item['responsible']['persons']:
                if responsible_item['person'] is not None:
                    responsible_text = responsible_item['person']['title']
                else:
                    responsible_text = '?'
                responsible_text += ' ' + responsible_item['service'] + ''
                responsible_list.append(responsible_text)

            responsible_text = ", ".join(responsible_list)
            document.add_paragraph(responsible_text)
            document.add_paragraph(item["note"])

            document.add_heading("TEMP Alle Item Informationen", level=3)
            document.add_paragraph(item.__str__())  # TODO #1 include serviceNotes

        document.save('output/' + agenda['name'] + '.docx')
