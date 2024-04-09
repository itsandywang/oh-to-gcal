from scraper import *

def main(event=None, context=None):
    scraper = Scraper()
    scraper.sync_assignments()
    scraper.sync_schedule()
    scraper.update_calendar()

if __name__ == '__main__':
    main()