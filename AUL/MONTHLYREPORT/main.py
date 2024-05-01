# standard library
import os
import argparse
import logging
import subprocess
# third party

# this project
import monthlyreport as mr


LOGGER = logging.getLogger()

## Example command on command prompt: pipenv run python run 2023 11 --root C:\Users\ylee\Desktop\dev\AUL\MONTHLYREPORT\etl

def main():
    args = get_args()
    action = args.action[0]
    year = None
    month = None
    if action == 'run':
        year = args.action[1]
        month = args.action[2]
    configure_logger(LOGGER, args.loglevel, args.log_server, year, month)
    
    root_dir = os.path.abspath(args.root)
    LOGGER.debug('parsed args.')

    rr = mr.ReportRunner(year,month,root_dir,args.server,args.database)
    if action == 'run':
        rr.run_all_claim_files()
    LOGGER.info('process complete')


def get_args():
    parser = argparse.ArgumentParser(description="Monthly reporting control program")
    parser.add_argument('action', nargs='+', help="Method to run, plus year and month if needed")
    parser.add_argument('--root', default='.', help="Root directory to find config files, default is current working directory")
    parser.add_argument('--server', default='AUL-DB-30', help="Target server, default is 'AUL-DB-30'")
    parser.add_argument('--database', default='AULDATAMART', help="Target database for insert, default is 'AULDATAMART'")
    parser.add_argument('--tableau_server', default='https://tableau.aulcorp.com', help="Target server, default is 'https://tableau.aulcorp.com'")
    parser.add_argument('--tableau_username', default='Yu-Sheng.Lee@protective.com', help="Target server, default is 'Yu-Sheng.Lee@protective.com'")
    parser.add_argument('--tableau_password', default='pw', help="Target server, default is 'pw'")
    parser.add_argument('--tableau_site', default='AUL-DB-30', help="Target server, default is 'AUL-DB-30'")
    parser.add_argument('--loglevel', default='INFO', help="Minimum level of log output (DEBUG/INFO/WARNING/ERROR/CRITICAL), default is 'INFO'")
    parser.add_argument('--log_server', nargs=3, help='Two-part log server details, expects "ip:port token"')
    parser.add_argument('--pause', dest='pause', action='store_true', help='Pause between steps, enter c to play remainder of instruction set without pauses')
    parser.set_defaults(pause=False)

    args = parser.parse_args()

    args.root = os.path.abspath(args.root)

    return args

def configure_logger(logger, level, log_server, year, month):
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s - %(message)s', '%Y-%m-%d %H:%M:%S')
    handler = logging.StreamHandler()
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    log_level = level.upper()
    logger.setLevel(log_level)
    if log_server:
        address, app = log_server
        url = "/api/{year}/q/{month}/post_log/{app_id}/SHORT/"
        url = url.format(year=year, month=month, app_id=app)
        httpHandler = logging.handlers.HTTPHandler(address, url)
        logger.addHandler(httpHandler)
        logger.info('Established connection to log server')

if __name__ == '__main__':
    main()
