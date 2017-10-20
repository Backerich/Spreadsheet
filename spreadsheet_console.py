class System(object):
    def exit(self):
        # Verlässt das Programm
        return sys.exit()

    def clear(self):
        # Räumt den Terminal Output auf
        os.system("cls" if os.name == "nt" else "clear")

    def continue_request(self, wb):
        # Abfrage ob Programm beendet werden soll oder weiter laufen soll
        try:
            self.continue_loop = input("\n" + strings_language[7] + "\n" + "> ").lower().strip()
        except KeyboardInterrupt:
            self.exit()

        if (self.continue_loop == "n") or (self.continue_loop == strings_language[0]):
            self.exit()

        else:
            # Wenn nicht 'exit' geht es wieder zum Ausgangsmenü zurück
            Menu().menu(wb)  # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! <- use this
            # self.clear.exit()
