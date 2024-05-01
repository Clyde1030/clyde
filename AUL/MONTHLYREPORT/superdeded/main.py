# standard library
import argparse
import logging
import logging.handlers
import os
import re
import sys
# this project
from monthlyreport import query, build_sql_headers



LOGGER = logging.getLogger()


def main():
    args = get_args()
    action = args.action[0] # should be a single verb for reset or [verb,year,quarter] for everything else
    year = None
    month = None
    if action in ('report'):
        year = args.action[1]
        month = args.action[2]


    root = r'C:\Users\ylee\Desktop\dev\AUL\MONTHLYREPORT'
    path = os.path.join(root,'etl','SQL Claims Detail with GWR_AULDATAMART_with Post Period_03-07-2020 (updated).sql')    
    intru_path = r"C:\Users\ylee\Desktop\dev\AUL\MONTHLYREPORT\etl\instructions.json"
    #"Large_Acct":"X:\\Dept.Risk.Management\\Large Accounts",
    #"External Reporting":"X:\\Dept.Risk.Management\\Monthly Reports\\External Reporting",


    configure_logger(LOGGER, args.loglevel, args.log_server, year, month)

    root_dir = os.path.abspath(args.root)
    # data_dir = args.data and re.sub(r'[\'"]$', '', args.data) # remove trailing single or double quotes (happens when directories are passed in as "X:\whatever\" including the trailing slash)
    LOGGER.debug('parsed args')
    # wrapper = monthlyreport.RsdbWrapper(year, quarter, root_dir, args.server, args.database,
    #                            data=data_dir, pause=args.pause, programs=programs_dir)

    if action == 'report':
        wrapper.report()
    else:
        wrapper.follow_instructions(action)
    LOGGER.info('process complete')



def get_args():
    parser = argparse.ArgumentParser(description="Monthly Reports Control Program")
    parser.add_argument('action', nargs='+', help="Method to run, plus year and month if needed")
    parser.add_argument('--root', default='.', help="Root directory to find config files, default is current working directory")
    parser.add_argument('--sqlserver', default='AUL-DB-30', help="Target server, default is 'localhost'")
    parser.add_argument('--database', default='AULDATAMART', help="Target database for insert, default is 'ReinsuranceSettlements'")
    parser.add_argument('--loglevel', default='INFO', help="Minimum level of log output (DEBUG/INFO/WARNING/ERROR/CRITICAL), default is 'INFO'")
    parser.add_argument('--log_server', nargs=3, help='Two-part log server details, expects "ip:port token"')
    parser.add_argument('--backup', default=None, help="Path for backup file")
    parser.add_argument('--tabserver', '-s', required=True, help='server address')
    parser.add_argument('--tabusername', '-u', required=True, help='username to sign into')
    parser.add_argument('--tabfilepath', '-f', required=True, help='filepath to the workbook')
    parser.add_argument('--tabsite', '-i', help='id of site to sign into')
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
