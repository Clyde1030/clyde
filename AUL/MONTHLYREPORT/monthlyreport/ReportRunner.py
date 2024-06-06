### This script combines Query.py to generate claim report according to instructions.json

# standard library
import json
import logging
import pathlib
import shutil
import os

# third party
import xlwings as xw

# this project
from .Query import query, build_sql_headers

CLAIM_TEMPLATE = r'X:\Dept.Risk.Management\Templates\Claims Detail_Template.xlsx'

LOGGER = logging.getLogger(__name__)

class ReportRunner:
    def __init__(self, year, month, root, server, database):
        LOGGER.debug("Creating ReportRunner class instance...")
        self.year = str(year) if year else ''
        self.month = str(month) if month else ''
        self.root = root
        self.working_dir = os.path.join(root,'etl')
        self.server = server
        self.database = database
        self.instructions_path = os.path.join(self.working_dir, 'instructions.json')
        self.instructions = self.get_instructions()
        self.sql_path = os.path.join(self.root,'etl','SQL Claims Detail with GWR_AULDATAMART_with Post Period_03-07-2020 (updated).sql')

    # Parse instructions.json and store the tasks into a dictionary that is available in the whole class
    # instructions.json in the etl folder has all the information for each report including the role id, file name, and location.
    def get_instructions(self, path=None):
        """return the instructions as a list of dictionary that includes all parameters for a report"""
        LOGGER.debug("Getting instructions from {}".format(path))
        if not path:
            path = self.instructions_path
        with open(path, 'r') as f:
            instruction_obj = json.load(f) # object_pairs_hook=collections.OrderedDict
        instructions = []

        for set_name in instruction_obj["instructions"]:
            for entity in instruction_obj["instructions"][set_name]:
                # Set up specifically for Coastal finance, which consists of different fc's
                if set_name == 'report' and (entity['name']=='Coastal Finance Company' or entity['name']=='Sensible Lending'):
                    for i in entity['id']:
                        empty_dict = {}
                        empty_dict['action'] = 'report'                 
                        empty_dict['name'] = entity['name']                  
                        empty_dict['role'] = 'fc'
                        empty_dict['id'] = i
                        empty_dict['claim_file_nm'] = entity['claim_file_nm']
                        empty_dict['tableau_file_nm'] = entity['tableau_file_nm']
                        empty_dict['resv_smry_file_nm'] = entity['resv_smry_file_nm']  
                        instructions.append(empty_dict)       
                elif set_name == 'report' and (entity['name']=='Maplewood Inver Grove Group'):
                    for i in entity['id']:
                        empty_dict = {}
                        empty_dict['action'] = 'report'                 
                        empty_dict['name'] = entity['name']                  
                        empty_dict['role'] = 'dlr_group'
                        empty_dict['id'] = i
                        empty_dict['claim_file_nm'] = entity['claim_file_nm']
                        empty_dict['tableau_file_nm'] = entity['tableau_file_nm']
                        empty_dict['resv_smry_file_nm'] = entity['resv_smry_file_nm']  
                        instructions.append(empty_dict)       
                elif set_name == 'report' and type(entity['id'])==int:
                    empty_dict = {}      
                    empty_dict['action'] = 'report'                   
                    empty_dict['name'] = entity['name']
                    empty_dict['role'] = entity['role']
                    empty_dict['id'] = entity['id']
                    empty_dict['claim_file_nm'] = entity['claim_file_nm']
                    empty_dict['tableau_file_nm'] = entity['tableau_file_nm']
                    empty_dict['resv_smry_file_nm'] = entity['resv_smry_file_nm']  
                    instructions.append(empty_dict)

        LOGGER.info("Finished adding instructions: {} counts to be processed".format(len(instructions)))
        LOGGER.debug(instructions)
        return instructions

    # Built on top of run_claim_file. Run all workbooks in the instructions.json
    def run_all_claim_files(self):
        filelist = []
        failure = []
        for i in self.instructions:
            if i['action'] == 'report':
                filelist.append(os.path.basename(i['claim_file_nm']))
                LOGGER.info('Running for id {}'.format(i['id']))
                try:
                    self.run_claim_file(i['claim_file_nm'],i['role'],i['id'],i['name'],'s')
                except Exception as e:
                    raise e    

    # Built on top of _query_ for. Paste the queried data to the assigned Excel workbook and formulize them properly 
    def run_claim_file(self,claim_file_path,role,id,name,mode):

        file_name = os.path.basename(claim_file_path)
        temp_file = os.path.join(self.working_dir,file_name)
        LOGGER.info('Processing {}'.format(file_name))

        # Copy previous month's workbook to etl folder. If not available, grab a new template
        try:
            if os.path.exists(temp_file)==False:
                shutil.copy2(claim_file_path,temp_file)
                LOGGER.debug("{} is copied to {}".format(file_name,self.working_dir))
        except FileNotFoundError:
            template_loc = pathlib.Path(CLAIM_TEMPLATE)
            shutil.copy2(template_loc,temp_file)
        except Exception as e:
            raise(e)

        # Get data from the query
        data, cols = self._query_for(role,id,mode)

        if data is not None:
            rowcount = data.shape[0]
            colcount = data.shape[1]
            data = data.values.tolist()

            with xw.App(visible=True) as Excel:
                # Open the workbook and append the data in a new data tab
                xw.App.display_alerts = False
                wb = xw.Book(temp_file)
                try:
                    wb.sheets.add('Query Result')
                except ValueError:
                    LOGGER.debug('Query Result tab already exists.')
                finally:
                    ws = wb.sheets['Query Result']
                    ws.clear_contents()
                    ws.range('A1').value = cols

                next_available_row = ws.range('A1').end('down').row+1
                if next_available_row == 1048577:
                    ws.range('A2').value = data        
                else:
                    ws.range('A'+str(next_available_row)).value = data        
                
                ws_data = wb.sheets['data']
                next_available_row = ws_data.range('A1').end('down').row+1
                if ws_data.range('AB1').value == 'GWR':
                    ws_data.range('A'+str(next_available_row)).value = data
                else:
                    range_to_copy = ws.range('A2:AA'+str(rowcount+1)).value
                    ws_data.range('A'+str(next_available_row)).value = range_to_copy                        
                
                # Copy down the formula and format in the data tab
                ws_data.range('A2:AU2').copy()
                ws_data.range('A3:'+'AU'+str(next_available_row+rowcount-1)).paste(paste='formats')
                ws_data.range('AC2:AU2').copy()
                ws_data.range('AC3:'+'AU'+str(next_available_row+rowcount-1)).paste(paste='formulas')


                # # Update the report date
                # pivot = wb.sheets['pivot']
                # pivot.range()


                # # Refresh all Pivot Tables within the workbook
                # wb.sheets['pivot'].select()
                # wb.api.ActiveSheet.PivotTables('PivotTable1').PivotCache().refresh()

                xw.App.display_alerts = True
                wb.save()

            LOGGER.info('{} is finished and saved at {}'.format(file_name,self.working_dir))
    
    # Rebuild the sql script with the new sql header of different modes. The output feeds into run_claim_file.
    def _query_for(self,role,id,mode):        
        '''
        role: 
            "fc" for finance company/lenders,\n
            "dlr" for dealer,\n
            "dlr_group" for dealer group\n
        mode default to s: 
            "s" for single month,\n
            "i" for ITD\n
        '''
        if mode == 's':
            with open(self.sql_path,'r') as f:
                qry_text = build_sql_headers(role,id,self.year,self.month,'s')+f.read()
        else:
            with open(self.sql_path,'r') as f:
                qry_text = build_sql_headers(role,id,self.year,self.month,'i')+f.read()            
        try:
            df, cols = query(qry_text,self.server,self.database)
        except UnboundLocalError:            
            LOGGER.info('0 record returned from query.')
            df = None
            cols = None
            pass                        

        return df, cols




# # # Debugging Section
# def configure_logger(logger, level, year, month):
#     formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s - %(message)s', '%Y-%m-%d %H:%M:%S')
#     handler = logging.StreamHandler()
#     handler.setFormatter(formatter)
#     logger.addHandler(handler)
#     log_level = level.upper()
#     logger.setLevel(log_level)

# def main():
#     root = r'C:\dev\AUL\MONTHLYREPORT'
#     # Large_Acct="X:\\Dept.Risk.Management\\Large Accounts"
#     # External_Reporting="X:\\Dept.Risk.Management\\Monthly Reports\\External Reporting"
#     configure_logger(LOGGER, 'info', 2022, 11)
#     rr = ReportRunner(2023,11,root,'AUL-DB-30','AULDATAMART')
#     rr.run_all_claim_files()
  
# if __name__=='__main__':
#     main()



