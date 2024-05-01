import logging
import os
import pprint
import tempfile
from collections import defaultdict
from functools import wraps

import tableauserverclient as TSC
from tableaudocumentapi import Workbook

LOGGER = logging.getLogger(__name__)

# """
# Submit thru command prompt:
# python publish_workbook.py -u usrname -s https://tableau.aulcorp.com -i site_id -f filename.twbx

# Then it would ask for password:
# Password: [enter the pw]
# """


def _sign_in_if_needed(func):
    @wraps(func)
    def wrapped_func(self, *args, **kwargs):
        if not self.conn:
            self._sign_in()
            new_connection = True
        else:
            new_connection = False

        res = func(self, *args, **kwargs)

        if new_connection and not self.keep_alive:
            self._close()
        
        return res
    return wrapped_func


class BaseTableau:
    def __init__(self, year, month, root, server, username, password, site_id):
        LOGGER.debug("Creating BaseTableau class instance...")
        self.year = str(year) if year else ''
        self.month = str(month) if month else ''
        self.root = root
        self.working_dir = os.path.join(root,'etl')
        self.server = server 
        self.username = username 
        self.password = password
        self.conn = None
        self.site_id = site_id        

    def __enter__(self):
        LOGGER.debug('Connecting BaseTableau ({})'.format(self.server))
        self.keep_alive = True
        return self
    
    def __exit__(self):
        LOGGER.debug('Exiting BaseTableau ({})'.format(self.server))
        self.keep_alive = False
        self._close()

    def _sign_in(self, username, password, site_id, target_server = None):
        s = target_server or self.server
        LOGGER.debug('Signing to Tableau server {}'.format(s))
        tableau_auth = TSC.TableauAuth(username, password, site_id)
        server = TSC.server(s)
        server.use_server_version() # Specify the server version
        server.auth.sign_in(tableau_auth)
        self.conn = server.auth 

    def _close(self):
        LOGGER.debug('Closing connection...')
        if self.conn:
            self.conn.sign_out()
        self.conn = None

    @_sign_in_if_needed
    def get_workbook_info(self, wbds):
        # Make a temp file for downloading the workbook
        server = self._sign_in(self.server, self.username, self.password, self.site_id)
        temp = tempfile.NamedTemporaryFile(delete=False)
        try:
            # Download the workbook into a temp file, without the extract
            server.workbooks.downloads(wb.id, temp.name, include_extract=False)
            # Open the workbook in the doc api and pull the info we need
            parsed = Workbook(temp.name)
            return parsed
        except Exception as e:
            print(e)
        finally:
            temp.close()
            os.remove(temp.name)     

    @_sign_in_if_needed
    def getpdf(self):

        server = self._sign_in(self.server, self.username, self.password, self.site_id)

        # Custom fields
        tag_to_filter = 'report'
        new_folder_path = 'Report/'

        with server.auth.sign_in(tableau_auth):
            # Specify a filter to only get workbooks tagged with 'report' tag
            tag_filter = TSC.RequestOptions()
            tag_filter.filter.add(TSC.Filter(TSC.RequestOptions.Field.Tags,
                                            TSC.RequestOptions.Operator.Equals,
                                            tag_to_filter))
            
            # Loop through filtered workbooks
            for workbook in TSC.Pager(server.workbooks, tag_filter):
                # Create a new directory for each workbook
                workbook_path = new_folder_path + workbook.name
                os.makedirs(workbook_path)

                # Get all views for workbook
                server.workbooks.populate_views(workbook)
            
                # Specifying PDF format (optional)
                size = TSC.PDFRequestOptions.PageType.A4
                orientation = TSC.PDFRequestOptions.Orientation.Landscape
                req_option = TSC.PDFRequestOptions(size, orientation)        

                # Loop through all views of a workbook
                for view in workbook.views:
                    # Get the PDF file from server
                    server.views.populate_pdf(view, req_option)


                # Save PDF file locally
                file_path = workbook_path + "/" + view.name + ".pdf"
                with open(file_path, "wb") as image_file:
                    image_file.write(view.pdf)
                
                print("\tPDF of {0} downloaded from {1} workbook".format(view.name, workbook))
        return


        # with server.auth.sign_in(tableau_auth):
        #     # Step 2: Get projects on the site, then look for the default one.
        #     all_projects, pagination_item = server.projects.get()
        #     default_project = next((project for project in all_projects if project.is_default()), None)
            
        #     # Step 3: If default project is found, continue with publishing 
        #     if default_project is not None:
        #         # Define publish mode, create new workbook, publish
        #         overwrite_true = TSC.Server.PublishMode.Overwrite
        #         new_workbook = TSC.WorkbookItem(default_project.id)
        #         new_workbook = server.workbooks.publish(new_workbook, args.filepath, overwrite_true)
        #         logging.debug("Workbook published. ID: {0}".format(new_workbook.id))
        #     else:
        #         error = "The default project could not be found."
        #         raise LookupError(error)

    def search(self):
        
        # This will hold the workbook names and fields in use
        # This could be a database or something else that stores this more permanently 
        USED_IN = defaultdict(list)
        DB_SERVER = 'mssql.test.tsi.lan'
        # The fun part! Loop over all workbooks on Server download them
        # grab their data sources and loop through those. If any of the data sources
        # connections point to the server I'm interested in, pull those fields out
        # and stick them into the USED_IN dictionary
        for wb in TSC.Pager(server.workbooks):
            workbook_info = get_workbook_info(wb)

            # Get data sources for the workbook
            datasources = workbook_info.datasources

            # Loop over the data sources and get connections
            for ds in datasources:
                connections = ds.connections

                # If it doesn't connect to the database we care about, ignore it.
                if not any(conn.server == DB_SERVER for conn in connections):
                    continue

                # For each field, check if it is in use.
                for f in ds.fields.values():
                    if f.worksheets:
                        USED_IN[wb.name].append((f.name, f.calculation))


        # Print out the results
        pprint.pprint(USED_IN)







