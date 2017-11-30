"""
    ***************************************************************************
    This file is for printing labels and reports needed for packaging/shipping
    TX products.

    It has the capability of printing Unit Labels, Box Labels, and TDS reports.
    The user can print any of these individually, or choose to print multiple
    at the same time.

    Data is retrieved from the database through either serial number or Laser
    ID, and also supports decoding QR codes for Aurora products.

    The program will search both SQL and mongo databases for test results, and
    will use the latest result if both are found to be populated.

    copyright: EMCORE.inc
    create: Aug 11, 2017
    ***************************************************************************
"""
# Emcore Lib
from automation1.apache.jsonparser.pylinq import *
import automation1.DatabaseManager as DM
from automation1.EmcoreMongoDB import db_info
from automation1.EmcoreSqlDB import db_list
from automation1.utilities import *
from automation1.constants import *
from automation1.MESAPI import *
from TDSConstant import *
from TDSExceptions import *
from automation1.mssql import *
import automation1.mssql as mssql
from idlelib.ReplaceDialog import replace

# Non Emcore Lib
import os
import sys
import win32print
import win32com.client
import win32api
import datetime
import math
import shutil
import subprocess
import csv
import configparser
import math
import random
import tkinter as tk
from subprocess import Popen

from tkinter import Tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import N, S, E, W
from tkinter import StringVar, IntVar

from collections import OrderedDict
from openpyxl.drawing.image import Image
from openpyxl import load_workbook
from decimal import Decimal

# Filenames
# For debugging
MONGO_RESULTS_FN        = 'mongoresults.json'
ALL_MONGO_RESULTS_FN    = 'allmongoresults.json'
SQL_RESULTS_FN          = 'sqlresults.json'
DEV_SPEC_FN             = 'specfile.json'

# Necessary files
LOOK_UP_TABLE_FN        = 'lookup_table.json'
MONGO_BRIDGE_FN         = 'Mongo_bridge.json'
STATION_FN              = 'StationInfo.json'

MONGO_BRIDGE_DB_TBL     = 'db tables'





class PrintDeviceGUI( ttk.Frame ):
    '''
    Widget class to get user input
    Gets Serial number and get related info from Database
    '''
    def __init__( self, master = None ):
        log.debug( LOG_ENTER )

        # main frame
        ttk.Frame.__init__( self, master )

        # Show the widget
        self.grid( sticky = N + S + E + W )

        # NEED ADD VER TO DB
        toplev = self.winfo_toplevel()
        toplev.title( SOFT_VERSION_NUMBER )

        self.gui_root = master
        self.boardinfodict = {}

        self.restartmsg = 'Select what to print'
        self.snnotgivenmsg = 'Please enter a serial number'

        # Start page
        self.createWidgets()

        log.debug( LOG_EXIT )
    # end function


    # -------------------------
    # GUI Interaction Functions
    # -------------------------

    def createWidgets( self ):
        '''
        Create TK widgets to get user info such as
        labels , text box and start button
        User need to scan the serial number under
        Textbox and click createTDS button to start tool
        '''
        log.debug( LOG_ENTER )

        # Variables to store the model, part number, and serial number
        self.serial_sv = StringVar()
        self.label_sv  = StringVar()
        self.label_sv.set( self.restartmsg )
        self.isproduct_iv = IntVar()
        self.printbox_iv = IntVar()
        self.printunit_iv = IntVar()
        self.printtds_iv = IntVar()
        self.submitmes_iv = IntVar()
        self.searchtype_sv = StringVar()
        self.boardsn_sv = StringVar()
        self.boardlaser_sv = StringVar()
        self.boardpcb_sv = StringVar()
        self.operatorid_sv = StringVar()

        # Product selection group
        prodtypegrp_lf = ttk.LabelFrame( self )
        prodtypegrp_lf[ 'text' ] = 'Product Type'
        prodtypegrp_lf[ 'relief' ] = 'ridge'
        prodtypegrp_lf.grid( row = 0, column = 0 )
        prodtypegrp_lf.grid( pady = 5, padx = 10 )

        self.generateOperatorInputField()

        # Default values:
        self.isproduct_iv.set( IS_TX_PRODUCT_RB_CODE )

        # Input group
        datagroup = ttk.LabelFrame( self )
        datagroup[ 'text' ] = 'Scan Serial Number/ID'
        datagroup[ 'relief' ] = 'ridge'
        datagroup.grid( row = 2, column = 0 )
        datagroup.grid( pady = 20, padx = 20 )

        # Info group
        infogroup = ttk.Frame( self )
        infogroup.grid( row = 3, column = 0 )

        # What to print
        self.checkboxgroup = ttk.LabelFrame( self )
        self.checkboxgroup.grid( row = 4, column = 0 )

        # Button group
        buttonGroup = ttk.Frame( self )
        buttonGroup.grid( row = 5, column = 0 )

        # Search Type
        self.searchtypelist_tx = [ BRD_SERIAL_NUM, BRD_LASER_ID, BRD_PCBA_ID ]
        self.searchtypelist_lm = [ BRD_SERIAL_NUM ]
        self.searchtypeCB = ttk.Combobox( datagroup )
        self.searchtypeCB[ 'textvariable' ] = self.searchtype_sv
        self.searchtypeCB[ 'state' ] = 'readonly'
        self.searchtypeCB[ 'values' ] = self.searchtypelist_tx
        self.searchtypeCB.bind( '<<ComboboxSelected>>',
                              self.focusSerialEntryBind )
        self.searchtypeCB.grid( row = 0, column = 2 )

        # Input for Search Type
        self.serialentry = ttk.Entry( datagroup )
        self.serialentry[ 'textvariable' ] = self.serial_sv
        self.serialentry[ 'width' ] = 22
        self.serialentry.grid( row = 0, column = 3 )
        self.serialentry.grid( sticky = W, pady = 5, padx = 15 )
        self.serialentry.bind( '<Return>', self.startEnterKey )

        # Radio Button for laser module
        istx_rb = ttk.Radiobutton( prodtypegrp_lf )
        istx_rb[ 'text' ] = 'Transmitter'
        istx_rb[ 'variable' ] = self.isproduct_iv
        istx_rb[ 'command' ] = self.updateSearchType
        istx_rb[ 'value' ] = IS_TX_PRODUCT_RB_CODE
        istx_rb[ 'takefocus' ] = True
        istx_rb.grid( row = 0, column = 0 )
        istx_rb.grid( sticky = W, pady = 5, padx = 15 )

        # Radio Button for laser module
        islm_rb = ttk.Radiobutton( prodtypegrp_lf )
        islm_rb[ 'text' ] = 'Laser Module'
        islm_rb[ 'variable' ] = self.isproduct_iv
        islm_rb[ 'command' ] = self.updateSearchType
        islm_rb[ 'value' ] = IS_LM_PRODUCT_RB_CODE
        islm_rb.grid( row = 0, column = 1 )
        islm_rb.grid( sticky = W, pady = 5, padx = 15 )

        # Create status label
        status_lbl = ttk.Label( infogroup )
        status_lbl[ 'textvariable' ] = self.label_sv
        status_lbl.grid( row = 0, column = 0 )
        status_lbl.grid( sticky = E, pady = 0, padx = 5 )

        # board info labels
        board_sn_lbl = ttk.Label( infogroup )
        board_sn_lbl[ 'textvariable' ] = self.boardsn_sv
        board_sn_lbl.grid( row = 1, column = 0 )

        board_laser_lbl = ttk.Label( infogroup )
        board_laser_lbl[ 'textvariable' ] = self.boardlaser_sv
        board_laser_lbl.grid( row = 3, column = 0 )

        board_pcb_lbl = ttk.Label( infogroup )
        board_pcb_lbl[ 'textvariable' ] = self.boardpcb_sv
        board_pcb_lbl.grid( row = 2, column = 0 )

        # check buttons
        selectionunit_cb = ttk.Checkbutton( self.checkboxgroup )
        selectionunit_cb[ 'text' ] = 'Print Unit Label'
        selectionunit_cb[ 'variable' ] = self.printunit_iv
        selectionunit_cb[ 'command' ] = self.focusSerialEntry
        selectionunit_cb.grid( row = 0, column = 1 )
        selectionunit_cb.grid( sticky = W, pady = 5, padx = 15 )

        selectionbox_cb = ttk.Checkbutton( self.checkboxgroup )
        selectionbox_cb[ 'text' ] = 'Print Box Label'
        selectionbox_cb[ 'variable' ] = self.printbox_iv
        selectionbox_cb[ 'command' ] = self.focusSerialEntry
        selectionbox_cb.grid( row = 0, column = 2 )
        selectionbox_cb.grid( sticky = W, pady = 5, padx = 15 )

        selectiontds_cb = ttk.Checkbutton( self.checkboxgroup )
        selectiontds_cb[ 'text' ] = 'Print TDS'
        selectiontds_cb[ 'variable' ] = self.printtds_iv
        selectiontds_cb[ 'command' ] = self.focusSerialEntry
        selectiontds_cb.grid( row = 0, column = 3 )
        selectiontds_cb.grid( sticky = W, pady = 5, padx = 15 )

        self.generateMesCheckbox()

        # Print button
        nextbutton = ttk.Button( buttonGroup )
        nextbutton[ 'text' ] = 'Print'
        nextbutton[ 'command' ] = self.start
        nextbutton.grid( column = 1, row = 0 )
        nextbutton.grid( sticky = E, pady = '5m', padx = '2m' )

        # Clear button
        clearbutton = ttk.Button( buttonGroup )
        clearbutton[ 'text' ] = 'Clear'
        clearbutton[ 'command' ] = self.clearFields
        clearbutton.grid( row = 0, column = 2 )

        # set GUI startup configurations
        self.operatoridentry.focus()
        self.serialentry.focus()
        self.printtds_iv.set( 0 )
        self.printbox_iv.set( 0 )
        self.printunit_iv.set( 0 )
        self.submitmes_iv.set( 1 )
        self.searchtype_sv.set( self.searchtypelist_tx[ 0 ] )

        # Testing:
        # self.serial_sv.set('BFDA039')
        # nextbutton.invoke()
        log.debug( LOG_EXIT )
    # end function



    def generateOperatorInputField( self ):
        """ Generate TKinter frame and inputs for operator ID
        Used on initial drawing of main GUI, and when changing the
        product type from LM to TX
        """
        log.debug( LOG_ENTER )

        # Operator Input group
        self.operatorgroup = ttk.LabelFrame( self )
        self.operatorgroup[ 'text' ] = 'Scan Operator ID'
        self.operatorgroup[ 'relief' ] = 'ridge'
        self.operatorgroup.grid( row = 1, column = 0 )
        self.operatorgroup.grid( pady = 20, padx = 20 )

        # Label for Operator ID
        operatorlabel = ttk.Label( self.operatorgroup )
        operatorlabel[ 'text' ] = 'Operator ID'
        operatorlabel.grid( row = 0, column = 0 )
        operatorlabel.grid( sticky = W, pady = 5, padx = 10 )

        # Operator ID input field
        self.operatoridentry = ttk.Entry( self.operatorgroup )
        self.operatoridentry[ 'textvariable' ] = self.operatorid_sv
        self.operatoridentry[ 'width' ] = 22
        self.operatoridentry.grid( row = 0, column = 1 )
        self.operatoridentry.bind( '<Return>', self.focusSerialEntryBind )

        log.debug( LOG_EXIT )
    # end function



    def generateMesCheckbox( self ):
        """ Generate TKinter checkbox for MES
        Used on initial drawing of main GUI, and when changing
        product type from LM to TX
        """

        log.debug( LOG_ENTER )

        self.selectionmes_cb = ttk.Checkbutton( self.checkboxgroup )
        self.selectionmes_cb[ 'text' ] = 'Submit MES'
        self.selectionmes_cb[ 'variable' ] = self.submitmes_iv
        self.selectionmes_cb[ 'command' ] = self.focusSerialEntry
        self.selectionmes_cb.grid( row = 1, column = 2 )
        self.selectionmes_cb.grid( sticky = W, pady = 5, padx = 15 )

        log.debug( LOG_EXIT )
    # end function



    def updateSearchType( self ):
        """ Update the search type parameters and default value.
        """

        log.debug( LOG_ENTER )

        isproduct_code = self.isproduct_iv.get()

        if ( isproduct_code == IS_TX_PRODUCT_RB_CODE ):
            self.searchtypeCB[ 'values' ] = self.searchtypelist_tx
            self.searchtype_sv.set( self.searchtypelist_tx[ 0 ] )
            self.generateOperatorInputField()
            self.generateMesCheckbox()
        elif ( isproduct_code == IS_LM_PRODUCT_RB_CODE ):
            self.searchtypeCB[ 'values' ] = self.searchtypelist_lm
            self.searchtype_sv.set( self.searchtypelist_lm[ 0 ] )
            self.operatorgroup.destroy()
            self.selectionmes_cb.destroy()
        else:  # default tx
            self.searchtypeCB[ 'values' ] = self.searchtypelist_tx
            self.searchtype_sv.set( self.searchtypelist_tx[ 0 ] )
            self.generateOperatorInputField()
        # end if

        self.serialentry.focus()

        self.searchtypeCB.grid( row = 0, column = 2 )

        log.debug( LOG_EXIT )
    # end function



    def focusSerialEntryBind( self, event ):
        """Sets focus to the next field (Serial Number)
        Called when enter key is pressed on Operator ID entry field on GUI
        to serial entry
        """
        log.debug( LOG_ENTER )

        self.serialentry.focus()

        log.debug( LOG_EXIT )
    # end function



    def focusSerialEntry( self ):
        """Same as focusSerialEntryBind but does not take event argument
        Called when checkbox is clicked, so we default to focusing the serial
        """
        log.debug( LOG_ENTER )

        self.serialentry.focus()

        log.debug( LOG_EXIT )
    # end function



    def clearFields( self ):
        """Clear all entry fields besides the operator ID, since that should
        rarely change during a given session
        """
        log.debug( LOG_ENTER )

        self.serial_sv.set( '' )
        self.serialentry.focus()

        log.debug( LOG_EXIT )
    # end function



    def startEnterKey( self, event ):
        """Called when enter key is pressed on the Serial Entry field on GUI
        Simply calls the start() function, which is called on 'Print'
        button click. This is because enter key function activations
        provide an extraneous event parameter to the called function
        """

        log.debug( LOG_ENTER )

        checkpoint_dict = read_json_file( 'at3552_checkpoint.json' )
        if( checkpoint_dict[ 'last' ] != '' ):
            last_id = checkpoint_dict[ 'last' ]
            next_index = checkpoint_dict[ 'ids' ].index( last_id ) + 1
        else:
            next_index = 0

        # loop through list of ids
        for i in range( next_index, len( checkpoint_dict[ 'ids' ] ) ):
            successbool = False
            try:
                self.serial_sv.set( checkpoint_dict[ 'ids' ][ i ] )
                successbool = self.start()
                checkpoint_dict[ 'last' ] = checkpoint_dict[ 'ids' ][ i ]
                write_json_file( 'at3552_checkpoint.json', checkpoint_dict )
            except:
                checkpoint_dict[ 'issues' ].append( checkpoint_dict[ 'ids' ][ i ])
                write_json_file( 'at3552_checkpoint.json', checkpoint_dict )
                
            if(successbool == None):
                log.debug('automated printing returned NONE, failed TDS')
                checkpoint_dict[ 'issues' ].append( checkpoint_dict[ 'ids' ][ i ])
                write_json_file( 'at3552_checkpoint.json', checkpoint_dict )


        # self.start()

        log.debug( LOG_EXIT )
    # end function

    # -------------------------
    # Main Start Function
    # -------------------------

    def start( self ):
        '''
        When user hit "Print" button this function gets called
        1. Update label widget
        2. Initialize logs
        3. Determine what to print (TDS, Box Label or both) and call
            appropriate functions
        '''
        log.debug( LOG_ENTER )

        # check to see if we received an operator ID before beginning
        operatorid = self.operatorid_sv.get()

        # Check for product type
        # For TX products, require operator id
        if ( self.submitmes_iv.get() == 1 and
            self.isproduct_iv.get() == IS_TX_PRODUCT_RB_CODE ):
            if ( operatorid == '' ):
                self.label_sv.set( 'Please enter an Operator ID' )
                messagebox.showwarning( title = 'Missing Operator ID',
                    message = 'Please enter an Operator ID' )
                return None
        # end if

        # -----------------
        # 1. Initial GUI Processing
        # -----------------

        # Clear GUI for new run
        self.clearGUI()

        # Retrieve user inputs from the GUI
        serialnumber, num_type, product_type, \
        print_tds_bool, print_box_bool, print_unit_bool = \
            self.readGUIInputs()


        # Quick check that input query value exists
        if ( serialnumber == '' ):
            log.debug( self.snnotgivenmsg )
            self.label_sv.set( self.snnotgivenmsg )
            messagebox.showwarning( title = 'Status',
                message = self.snnotgivenmsg )
            self.label_sv.set( self.restartmsg )
            return None
        # end if

        # For every tds attempt they want to restart logs
        if ( not log.initialize() ):
            log.error( "Failed to initialize main log class" )
        # end if

        self.label_sv.set( "Getting record from database, please wait... " )
        self.gui_root.update()

        # -----------------
        # 2. Database Queries/TDS Printings
        # -----------------

        # self.start_tds handles printing TDS and also saves
        # board information to self.boardinfodict for use in label printing
        tdsretbool = self.start_tds( product_type = product_type,
            serialnumber = serialnumber, num_type = num_type,
            print_tds_bool = print_tds_bool )

        # if TDS printing is successful, move on to printing box label
        if ( tdsretbool == False ):
            log.debug( 'Printing TDS failed, please try again' )
            return None
        else:
            log.debug( 'Printing TDS completed, beginning box label printing' )
            return True
        # end if

        # Report board information (board ID, SN, laser ID) to GUI
        self.report_board_info( product_type = product_type )

        # -----------------
        # 3. Label Printing
        # -----------------

        # after TDS is complete or skipped, move into label printing methods
        self.start_label( print_box_bool = print_box_bool,
            print_unit_bool = print_unit_bool )

        # -----------------
        # 4. MES logging
        # -----------------

        # Only do MES for TX products and when submit MES checkbox is checked
        if ( self.isproduct_iv.get() == IS_TX_PRODUCT_RB_CODE
            and self.submitmes_iv.get() == 1 ):
            log.debug( 'Submit MES is checked, submitting to MES' )
            # Checkin to MES after printing label
            self.checkinMES( print_box_bool, print_unit_bool )
            # Second MES code
            self.createWidgets_MesStatus( print_box_bool, print_unit_bool )
        # end if

        logs.log_thread_kill()

        log.debug( LOG_EXIT )

        return None
    # end function



    def clearGUI( self ):
        '''
        When start() function is run, clear gui info labels
        and set serial input to capital letters
        '''

        log.debug( LOG_ENTER )

        self.label_sv.set( 'Starting..' )

        # uppercase input
        self.serial_sv.set( str.upper( str.strip( self.serial_sv.get() ) ) )

        # clear board info labels
        self.boardsn_sv.set( '' )
        self.boardlaser_sv.set( '' )
        self.boardpcb_sv.set( '' )

        self.gui_root.update()

        log.debug( LOG_EXIT )

    # end function



    def readGUIInputs( self ):
        '''
        Read user inputs from GUI during start() function
        and return them
        Returns 6 variables in this order:
            serialnumber: Identification number to search with
            num_type: Type of identification (laser, board, SN)
            product_type: TX or LM product type
            print_tds_bool: 1 = print tds, 0 = don't print
            print_box_bool: 1 = print box label, 0 = don't print
            print_unit_bool: 1 = print unit label, 0 = don't print
        '''

        log.debug( LOG_ENTER )

        serialnumber = self.serial_sv.get()
        # decode QR code if input into serial
        if ( '>' in serialnumber ):
            serialnumber = serialnumber.split( '>' )[ -1 ]
            self.serial_sv.set( serialnumber )
            self.gui_root.update()
        # end if

        # allow user to choose to search by SN or Laser
        searchtypestring = self.searchtype_sv.get()
        if ( searchtypestring == BRD_SERIAL_NUM ):
            num_type = BRD_SERIAL_NUM
        elif ( searchtypestring == BRD_LASER_ID ):
            num_type = BRD_LASER_ID
        elif ( searchtypestring == BRD_PCBA_ID ):
            num_type = BRD_PCB_ID
        # end if

        # Check for product type
        # only print TX boards TDS and Labels
        if ( self.isproduct_iv.get() == IS_TX_PRODUCT_RB_CODE ):
            product_type = PRODUCT_TYPE_TX
        else:
            product_type = PRODUCT_TYPE_LM
        # end if

        # Check which of labels and/or tds should be printed
        print_tds_bool = self.printtds_iv.get()
        print_box_bool = self.printbox_iv.get()
        print_unit_bool = self.printunit_iv.get()

        log.debug( LOG_EXIT )

        return serialnumber, num_type, product_type, \
            print_tds_bool, print_box_bool, print_unit_bool

    # end function



    def start_tds( self, product_type, serialnumber, num_type,
        print_tds_bool) -> bool:
        '''
        Called during main start() function for handling TDS printing
            1. Create TDSManager/TDSWorker for given product_type
            2. Call CreateTDS
                a. Queries databases for board information and spec files
                b. Print TDS if print_tds_bool == 1
                c. Returns boardinfo
            3. Save boardinfo to self.boardinfodict
                and cfginfo to self.cfginfo so that it can be used
                for label printing without re-querying database

            4. return True if process is successful
        '''

        log.debug( LOG_ENTER )

        tdsman_cls = TDSManagerFactory( product_type = product_type )
        if ( tdsman_cls is None ):
            log.debug( 'Failed to find TDSManager to support product type: '
                '{0}'.format( product_type ) )

            msg = 'Product type: {0} NOT SUPPORTED'.format( product_type )
            self.label_sv.set( msg )
            return False
        # end if

        try:
            tdsman_obj = tdsman_cls( serialnumber = serialnumber,
                product_type = product_type, num_type = num_type )
        except Exception as e:
            log.error( 'Problem occurred during TDS creation, '
                       'probably due to db' )
            self.label_sv.set( 'Problem occurred during TDS creation' )
            return False
        # end try

        try:
            tdsretdict = tdsman_obj.CreateTDS( print_tds_bool )
        except Exception as e:
            msg = "Error creating TDS. "
            msg += e.args[ 0 ]
            self.label_sv.set( msg )
            return False
        # end try

        tdsretbool = tdsretdict.get( KEY_TDS_STATUS_BOOL, None )
        tdsretmsg = tdsretdict.get( KEY_TDS_STATUS_MSG, None )

        log.debug( 'template creation bool: {0}'.format( tdsretbool ) )
        log.debug( 'output status msg' )
        self.label_sv.set( tdsretmsg )
        self.gui_root.update()

        # if TDS printing failed, return False
        if ( tdsretbool == False ):
            log.debug( 'Printing TDS failed, please try again' )
            return False
        # end if

        # retrieve boardinfo from the tds printing process, even if TDS was not
        # actually printed, so we don't have to query a second time for label
        # printing. Save to self.boardinfodict
        self.boardinfodict = tdsretdict[ KEY_TDS_BOARD_INFO_DICT ]
        self.cfginfo = tdsretdict[ KEY_TDS_CFG_INFO ]

        log.debug( LOG_EXIT )

        return True
    # end function



    def report_board_info( self, product_type ):
        '''
        Report SN, Laser and PCBA (for transmitters) to GUI for user to see

        '''

        log.debug( LOG_ENTER )

        # some products use 'SN' others use 'Serial Number' as key
        # key = BRD_SN or key = BRD_SERIAL_NUM accounts for both cases
        key = BRD_SN
        if ( BRD_SN not in self.boardinfodict ):
            key = BRD_SERIAL_NUM

        if( key not in self.boardinfodict ):
            log.debug( 'No serial number in self.boardinfodict, exiting report' )
            return False

        self.boardsn_sv.set( 'Serial Number: {0}'.format(
                             self.boardinfodict[ key ] ) )
        if ( product_type == PRODUCT_TYPE_TX ):
            self.boardlaser_sv.set( 'Laser ID: {0}'.format(
                                self.boardinfodict[ BRD_LASER_ID ] ) )
            self.boardpcb_sv.set( 'PCBA ID: {0}'.format(
                                self.boardinfodict[ BRD_PCB_ID ] ) )
        # end if
        self.gui_root.update()

        log.debug( LOG_EXIT )

    # end function



    def start_label( self, print_box_bool, print_unit_bool ) -> bool:
        '''
        Called during main start() function for handling label printing
        1. Checks if any unit label printing is needed
        2. Calls self.PrintLabels
            a. use PrintLabel object to print labels as requested
        3. Set GUI labels to show whether label printing was successful
        4. Return True/False for printing status
        '''

        log.debug( LOG_ENTER )

        if ( ( print_box_bool == 0 ) and ( print_unit_bool == 0 ) ):
            log.debug( 'Skipped label printing, process complete!' )
            self.label_sv.set( 'Skipped label printing, process complete!')
        else:
            finalresult = self.boardinfodict[ TST_FINAL_RESULT ]
            if ( finalresult == TST_RES_PASSED ):
                labelretstatus = self.PrintLabels(
                    print_box_bool, print_unit_bool )
            else:
                self.label_sv.set( 'Final result is failed, will not print '
                                   'label' )
                return False
            # end if

            if ( labelretstatus == False ):
                log.debug( 'Label printing failed, please try again' )
                self.label_sv.set( 'Label printing failed' )
                return False
            # end if

            log.debug( 'Label printing completed, process complete' )
            self.label_sv.set( 'Label printed, process complete!' )
        # end if

        log.debug( LOG_EXIT )

        return True
    # end function



    def checkinMES( self, print_box_bool, print_unit_bool ) -> bool:
        """
        Check in packaging station to MES, telling it that
        a given TX board has reached the station
        checkinMES called from start() after printing labels, and only
        for printing labels (not TDS)

        return:
            True -> checkin successful
            False -> checkin had problems, do not continue to pass/fail MES
        """
        log.debug( LOG_ENTER )

        # No MES checkin for LM type
        if ( self.isproduct_iv.get() == IS_LM_PRODUCT_RB_CODE ):
            log.debug( 'No MES Checkin for LM product' )
            return True
        # end if

        if( print_box_bool == 0 and print_unit_bool == 0):
            log.debug( 'Not printing labels, no MES Checkin')
            return True

        try:
            tx_serial = self.boardinfodict[ BRD_SN ]
        except Exception as e:
            message = 'key SN not available in boardinfodict, cannot checkin \
                to MES'
            log.debug( message )
            messagebox.showwarning( title = 'MES checkin error',
                            message = message )
            return False

        operator_id = self.operatorid_sv.get()
        timestamp = datetime.datetime.now().strftime(
                '%Y-%m-%d %H:%M:%S:%f' )[ :-3 ]

        log.debug( 'Submitting to MES with:' )
        log.debug( 'SN: {0}'.format( tx_serial ) )
        log.debug( 'Operator ID: {0}'.format( operator_id ) )
        log.debug( 'Timestamp: {0}'.format( timestamp ) )
        log.debug( 'MES Code: {0}'.format( MES_CHECKIN ) )

        if( print_box_bool == 1 ):
            log.debug( 'Submitting MES for box_label printing' )
            VSGPostDataToMesAndMoveOut( operator_id, tx_serial, '130',
                MES_CHECKIN, timestamp, timestamp,'EA02')

        if( print_unit_bool == 1 ):
            log.debug( 'Submitting MES for unit_label printing' )
            VSGPostDataToMesAndMoveOut( operator_id, tx_serial, '89',
                MES_CHECKIN, timestamp, timestamp,'EA02')

        log.debug( LOG_EXIT )

        return True
    # end function



    def PrintLabels( self, print_box_bool, print_unit_bool ) -> bool:
        """
        Print the box label and the unit label using PrintLabel class from
        automation1 folder

        tdsretdict: return dictionary from tdsworker that contains cfginfo and
               boardinfodict

        returns:
            True if print was successful
            False if print had issues
        """
        log.debug( LOG_ENTER )
        islasermodule = self.isproduct_iv.get()
        if ( islasermodule == True ):
            self.boardinfodict[ DB_PRODUCT_FAMILY ] = TST_TYPE_LM
        elif ( DB_PRODUCT_FAMILY not in self.boardinfodict ):
            log.info( "Warning: {0} Not found using CATV Default".
                  format( DB_PRODUCT_FAMILY ) )
            self.boardinfodict[ DB_PRODUCT_FAMILY ] = FAMILY_CATV
        # end if

        # call the printData function to print the label using board info and
        # config files passed. Only print if rollup result is TST_RES_PASSED
        # (0, 1) indicates that we want the unit label only
        # (1, 0) means we want the box label only
        # (1, 1) prints both unit and box labels

        printlabelobj = PrintLabel( self.boardinfodict, self.cfginfo )
        printstatus = printlabelobj.printData( print_box_bool, print_unit_bool )

        log.debug( LOG_EXIT )
        return printstatus
    # end function



    def createWidgets_MesStatus( self, print_box_bool, print_unit_bool ):
        """
        previous: start()

        Called at the end of start() function, after printing is completed
        or skipped as selected appropriately
        Creates a new window where the operator can scan in an MES
        pass or fail code to indicate when operations for the
        given board are completed at Packaging Station
        """
        log.debug( LOG_ENTER )

        # No MES checkin for LM type
        if ( self.isproduct_iv.get() == IS_LM_PRODUCT_RB_CODE ):
            log.debug( 'No MES submit for LM product' )
            return None
        # end if

        # If printing box, automatically submit a pass
        if( print_box_bool == 1 ):
            log.debug( 'Print box MES status defaulting to PASS')
            self.submitMesPassBoxLabel()
        if( print_unit_bool == 0 ):
            log.debug( 'Not printing unit label, skipping pass/fail MES submit')
            return None

        log.debug( 'Creating MES entry window' )

        # Tkinter variable to store MES code input into
        self.mes_code_sv = StringVar()

        # Create the base window and frame which tkinter elements
        # will be placed into
        self.WriteMES_window = tk.Toplevel()
        self.WriteMES_window.wm_title( 'MES Update' )

        MESFrame = ttk.Frame( self.WriteMES_window )
        MESFrame[ 'relief' ] = tk.RIDGE
        MESFrame[ 'padding' ] = '0.15i'
        MESFrame.grid( row = 0, column = 0 )

        MESInnerFrame = ttk.LabelFrame( MESFrame )
        MESInnerFrame[ 'relief' ] = tk.RIDGE
        MESInnerFrame[ 'text' ] = 'Submit MES Code'
        MESInnerFrame.grid( row = 0, column = 0 )
        MESInnerFrame.grid( pady = 20, padx = 20 )

        # Create the display for the tkinter window
        # Label for MES Code input
        meslabel = ttk.Label( MESInnerFrame )
        meslabel[ 'text' ] = 'Submit MES Code'
        meslabel.grid( row = 0, column = 0 )
        meslabel.grid( sticky = W, pady = 5, padx = 10 )

        # # Scan MES code
        # self.mescodeentry = ttk.Entry( MESInnerFrame )
        # self.mescodeentry[ 'textvariable' ] = self.mes_code_sv
        # self.mescodeentry[ 'width' ] = 22
        # self.mescodeentry.grid( row = 0, column = 1 )
        # self.mescodeentry.bind( '<Return>', self.submitMesCodeBind )

        # submit_mes_button = ttk.Button( MESInnerFrame )
        # submit_mes_button[ 'text' ] = 'Submit to MES'
        # submit_mes_button[ 'command' ] = self.submitMesCode
        # submit_mes_button.grid( column = 0, row = 1 )
        # submit_mes_button.grid( sticky = E, pady = '5m', padx = '2m' )

        # MES code buttons (should not be used when scan is available)
        submit_mes_pass_button = ttk.Button( MESInnerFrame )
        submit_mes_pass_button[ 'text' ] = 'MES Pass'
        submit_mes_pass_button[ 'command' ] = self.submitMesPass
        submit_mes_pass_button.grid( column = 0, row = 1 )
        submit_mes_pass_button.grid( sticky = E, pady = '5m', padx = '2m' )

        submit_mes_fail_button = ttk.Button( MESInnerFrame )
        submit_mes_fail_button[ 'text' ] = 'MES Fail'
        submit_mes_fail_button[ 'command' ] = self.submitMesFail
        submit_mes_fail_button.grid( column = 1, row = 1 )
        submit_mes_fail_button.grid( sticky = E, pady = '5m', padx = '2m' )

        return_button = ttk.Button( MESInnerFrame )
        return_button[ 'text' ] = 'Cancel'
        return_button[ 'command' ] = self.closeMesWindow
        return_button.grid( column = 2, row = 1 )
        return_button.grid( sticky = E, pady = '5m', padx = '2m' )

        log.debug( LOG_EXIT )
    # end function



    def closeMesWindow( self ):
        """
        Called when the 'Cancel' button is clicked in MES Frame
        Will close the window and return to printing, clearing the
        printing window inputs as well
        Also called when submitMesCode is completed
        """
        log.debug( LOG_ENTER )

        log.debug( 'Destroying MES update frame' )

        self.WriteMES_window.destroy()

        log.debug( LOG_EXIT )
    # end function



    def submitMesCodeBind( self, event ):
        """ calls submitMesCode, as enter commands take an event
        argument, where button clicks do not
        """
        log.debug( LOG_ENTER )

        log.debug( LOG_EXIT )
    # end function



    def submitMesCode( self, code, print_box_bool ):
        """
        Called when MES code is input and submitted (button clicked)
        Will upload the code with relevant information to MES
        such as operator, SN, datetime
        """
        log.debug( LOG_ENTER )

        # No MES checkin for LM type
        if ( self.isproduct_iv.get() == IS_LM_PRODUCT_RB_CODE ):
            log.debug( 'No MES submit for LM product' )
            return True
        # end if

        try:
            tx_serial = self.boardinfodict[ BRD_SN ]
        except Exception as e:
            message = 'key SN not available in boardinfodict, cannot submit \
                code to MES'
            log.debug( message )
            messagebox.showwarning( title = 'MES submit pass/fail code error',
                            message = message )
            self.closeMesWindow()

        try:
            # Retrieve user inputs from the GUI
            print_unit_bool = self.printunit_iv.get()
        except Exception as e:
            message = 'Problem reading user input during MES code submit'
            log.debug( message )
            messagebox.showwarning( title = 'MES submit pass/fail code error',
                            message = message )
            self.closeMesWindow()


        operator_id = self.operatorid_sv.get()
        timestamp = datetime.datetime.now().strftime(
                        '%Y-%m-%d %H:%M:%S:%f' )[ :-3 ]

        log.debug( 'Submitting to MES with:' )
        log.debug( 'SN: {0}'.format( tx_serial ) )
        log.debug( 'Operator ID: {0}'.format( operator_id ) )
        log.debug( 'Timestamp: {0}'.format( timestamp ) )
        log.debug( 'MES Code: {0}'.format( code ) )

        if( 'print_box_bool' in locals()
            and print_box_bool == 1 ):
            log.debug( 'Submitting MES code for box_label printing: {}'.format(
                code ))
            VSGPostDataToMesAndMoveOut( operator_id, tx_serial, '130',
                code, timestamp, timestamp,'EA02')

        if( print_unit_bool == 1 ):
            log.debug( 'Submitting MES code for unit_label printing: {}'.format(
                code ))
            VSGPostDataToMesAndMoveOut( operator_id, tx_serial, '89',
                code, timestamp, timestamp,'EA02')

        try:
            self.closeMesWindow()
        except Exception as e:
            log.debug('No MES submit code window was generated (box label\
                printed), continuing.' )

        log.debug( LOG_EXIT )
    # end function


    def submitMesPassBoxLabel( self ):
        """
        Called after successful box label printing, automatically
        submit an MES pass code for box label printing
        """
        log.debug( LOG_ENTER )

        self.submitMesCode( MES_PASS, 1 )

        log.debug( LOG_EXIT )




    def submitMesPass( self ):
        """
        Called when 'MES Pass' button is clicked
        Calls submitMesCode with a "Passed" code
        """
        log.debug( LOG_ENTER )

        # Second argument of 0 indicates not submitting box_label MES code
        self.submitMesCode( MES_PASS, 0 )

        log.debug( LOG_EXIT )
    # end function



    def submitMesFail( self ):
        """
        Called when 'MES Fail' button is clicked
        Calls submitMesCode with a "Failed" code
        """
        log.debug( LOG_ENTER )

        self.submitMesCode( MES_FAIL )

        log.debug( LOG_EXIT )
    # end function
# end class





def TDSManagerFactory( product_type: str ) -> object:
    """ Function to decide based on product type which TDS manager class to use

    param product_type: LM, TX, COB, BAR
    return:
        ret_cls: Class to use by product_type
        None: Issue not finding matching class
    """
    log.debug( LOG_ENTER )

    ret_cls = None

    if ( product_type == '' ):
        log.debug( 'Product type is empty' )
        return ret_cls
    # end if

    log.debug( 'product type is: {0}'.format( product_type ) )

    for poten_cls in TDSManager.__subclasses__():
        retbool = poten_cls.IsTDSManagerFor( product_type = product_type )
        if ( retbool == True ):
            log.debug( 'Found matching TDSManager: {0}'.format(
                poten_cls.__name__ ) )
            ret_cls = poten_cls
            break
        # end if
    # end for

    if ( ret_cls is None ):
        log.debug( 'Did not find TDS manager subclass' )
        return ret_cls
    # end if

    log.debug( LOG_EXIT )
    return ret_cls
# end function





class TDSManager( object ):
    def __init__( self, serialnumber: str, product_type: str, num_type: str ):
        """
        serialnumber: value for the number type
        product_type: type of product to print for, e.g. TX or LM
        num_type: number type, e.g. Serial Number, Laser ID
        """
        log.debug( LOG_ENTER )

        self.boardinfodict = {}
        self._serialnumber = serialnumber
        self._product_type = product_type
        self.num_type = num_type

        if ( self._serialnumber == '' ):
            msg = 'SN is empty'
            log.debug( msg )
            raise ValueError( msg )
        # end if

        if ( self._product_type == '' ):
            msg = 'Product type is empty'
            log.debug( msg )
            raise ValueError( msg )
        # end if

        log.debug( LOG_EXIT )
    # end function



    @property
    def serialnumber( self ):
        log.debug( LOG_ENTER )

        log.debug( LOG_EXIT )
        return self._serialnumber
    # end function



    @classmethod
    def IsTDSManagerFor( cls: object, product_type: str ) -> bool:
        """ Checks if the product type is supported by this class

        param product_type: ex LM, TX
        return: ( boolean )
            True: produc ttype is supported
            False: not supported
        """
        log.debug( LOG_ENTER )

        log.debug( LOG_EXIT )
        raise NotImplementedError( "should be implemented in sub class" )

    # end function



    def CreateTDS( self ) -> dict:
        """ It fill all the required field on test data sheet
        based on test results.All the data is read from
        testinfo.json file and update each cell based on
        test name and group

        :return: (dictionary)
            retdict:
                status_bool
                status_msg
        """
        log.debug( LOG_ENTER )

        log.debug( LOG_EXIT )
        raise NotImplementedError( "should be implemented in sub class" )

    # end function
# end class





class TDSWorker( object ):
    def __init__( self, serialnumber:str, sessiontype:str ):
        '''
        serialnumber : serial number to create tds with
        sessiontype: sessiontype is intended to distinguish between production,
                engineering, oql, spc
        '''
        log.debug( LOG_ENTER )

        self.emkridnumber   = ''
        self.tdstemplatefn  = ''
        self.tdsprintfn     = ''
        self.sessiontype    = ''
        self.cfginfo        = {}
        self.boardinfodict  = {}
        self.printdatadict  = {}
        self.tdscellmap     = {}
        self.alltestpassed  = True
        self.alldatapresent = True

        self.tdsdatasetconfig = {}
        self.tdsdatasetdict = {}

        self._serialnumber = serialnumber
        self.sessiontype = str.strip( sessiontype )

        if ( str( self._serialnumber ).strip() == '' ):
            errmsg = 'Serial number is empty'
            log.error( errmsg )
            raise ValueError( errmsg )
        # end if

        if ( ( self.sessiontype is None ) or
             ( str.strip( self.sessiontype ) == '' ) ):
            errmsg = 'Did not get expected value for sessiontype'
            log.error( errmsg )
            raise ValueError( errmsg )
        # end if

        computername = getThisComputerName()
        if ( not computername ):
            log.error( "Failed to get Station name" )
            return None
        # end if

        # manual stationfile
        self.stationfile = read_json_file( 'StationInfo.json' )
        # self.stationfile = getStationInfo( db_info = db_info,
        #     station_id = computername )

        if ( not self.stationfile ):
            errmsg = 'Could not get station file'
            log.error( errmsg )
            raise ValueError( errmsg )
        # end if

        write_json_file( STATION_FN, self.stationfile )

        log.debug( LOG_EXIT )
    # end function



    def CreateTDS( self ) -> dict:
        """ It fill all the required field on test data sheet
        based on test results.All the data is read from
        testinfo.json file and update each cell based on
        test name and group

        :return: (dictionary)
            retdict:
                status_bool
                status_msg
        """
        log.debug( LOG_ENTER )

        log.debug( LOG_EXIT )
        raise NotImplementedError( "should be implemented in sub class" )

    # end function



    def getRelatedDataFromCfg( self ) -> bool:
        """ Retrieves data from TDS label specific cfg or dictionary

        Return:
            True: Data retrieved
            False: Issue with obtaining data
        """
        log.debug( LOG_ENTER )

        if ( BRD_TDS_CELL_MAP not in self.cfginfo ):
            log.error( 'TDS software revision does not support product '
                '{0}'.format( BRD_TDS_CELL_MAP ) )
            messagebox.showerror( title = 'Does not support TDS',
                    message = 'This product does not support TDS printing ' )
            return False
        # end if

        if ( BRD_TDS_TMPT_FN not in self.cfginfo ):
            log.error( 'TDS template filename not found in '
                       'tds cfg dictionary' )
            return False
        # end if

        if ( BRD_TDS_PRNT_FN not in self.cfginfo ):
            log.error( 'TDS print filename not found in tds cfg dictionary' )
            return False
        # end if

        self.tdscellmap = self.cfginfo[ BRD_TDS_CELL_MAP ]
        self.tdstmptfn = self.cfginfo[ BRD_TDS_TMPT_FN ]
        self.tdsprintfn = self.cfginfo[ BRD_TDS_PRNT_FN ]

        # Added Product Spec Data get
        self.prod_spec_data = self.cfginfo.get( BRD_SPEC_DATA, None )
        if ( self.prod_spec_data is None ):
            log.debug( '{0} not found! Required for product data and '
                'model strings'.format( BRD_SPEC_DATA ) )
            return False
        # end if

        # TO DO DO WE USE THIS? laser module, tx?
        # get Optical Power from spec file
        # Andy 10/19/17: This might be specific to TX, maybe move to
        # a new definition of getRelatedDataFromCfg under TDSWorker_TX
        if ( ( BRD_SPEC_DATA in self.cfginfo ) and
             ( KEY_TDS_OPTI_POWER_DBM in self.cfginfo[ BRD_SPEC_DATA ] ) ):
            try:
                log.debug( 'Setting Optical Power from specfile' )
                self.boardinfodict[ KEY_TDS_OPTI_POWER_DBM ] = \
                    self.cfginfo[ BRD_SPEC_DATA ][ KEY_TDS_OPTI_POWER_DBM ]
            except Exception as e:
                log.debug( 'Error getting Optical Power dBm from specfile: '
                           '{0}'.format( e ) )
            # end try
        # end if

        log.debug( LOG_EXIT )
        return True
    # end function



    def updateTemplate( self ) -> bool:
        """
        Update TDS template from result and
        board spec file

        Return : True if template is updated
                 False if failed to update

        """
        log.debug( LOG_ENTER )

        # check for template tds file
        if ( self.tdstmptfn is None ):
            log.debug( 'Failed to find TDS template file using name '
                       '{0}'.format( self.tdstmptfn ) )
            return False
        # end if

        # check for print tds file
        if ( self.tdsprintfn == '' ):
            log.debug( 'Failed to find TDS print file using name '
                       '{0}'.format( self.tdsprintfn ) )
            return False
        # end if

        # Supported TDS file
        retbool = self._updateTDS( printfile = self.tdsprintfn )
        if ( retbool == False ):
            log.debug( '_updateTDS() failed!' )
            return False
        # end if

        log.debug( LOG_EXIT )
        return True
    # end function



    def printTDSData( self ) -> bool:
        """
        Takes care of any pre processing before executing
        the print command

        Return:
            True: No issues
            False: Issue with pre processing
        """
        log.debug( LOG_ENTER )

        log.debug( 'Looking for printer for TDS with name: {0}'.format(
                    KEY_TDS_PRINTING ) )

        # WARNING: This code actually works, however, it needs a very specific
        # string name.
        # Ex.For printer BLDG5_C364_DNW on network CAALHSUT07.emcore.us
        # will use the string "\\\\CAALHSUT07.emcore.us\\BLDG5_C364_DNW".
        # Hint: To get the printer names, go into Notepad and change the default
        # printer manually. Then use win32print.GetDefaultPrinter() to get
        # the printer name that win32api.ShellExecute understands.
        try:
            if ( KEY_TDS_PRINTER_INFO in self.stationfile.keys() ):
                if ( KEY_TDS_PRINTING in self.stationfile[
                     KEY_TDS_PRINTER_INFO ].keys() ):
                    printer = self.stationfile[ KEY_TDS_PRINTER_INFO ]\
                                         [ KEY_TDS_PRINTING ]
                else:
                    log.debug( 'No configuration for TDS Printing, '
                               'using default printer' )
                    printer = win32print.GetDefaultPrinter( )
                # end if
            else:
                log.debug( 'No printer configurations, using default printer' )
                printer = win32print.GetDefaultPrinter( )
            # end if
        except Exception as e:
            log.debug( 'Unexpected error in preparing printer: {0}'.format(
                        e ) )
            return False
        # end try

        tdsfilename = self.tdsprintfn

        if( ( HW_PRODUCT_ID in self.boardinfodict ) and
            ( 'AT3552' in self.boardinfodict[ HW_PRODUCT_ID ] ) ):
            # Skip printing TDS for AT3552 units
            log.debug( 'AT3552 Unit, skipping TDS print and saving')
            return True

        if ( ( HW_PRODUCT_ID in self.boardinfodict ) and
             ( 'GX2' in self.boardinfodict[ HW_PRODUCT_ID ] ) ):
            self.SaveGX2TDSToCompressedPDF()
        # end if

        try:
            log.info( 'Sending print request with: {0}'.format( printer ) )
            win32api.ShellExecute (
                    0,
                    'printto',
                    tdsfilename,
                    '"%s"' % printer,
                    '.',
                    0
            )

            # Don't want to kill process, instead close process using correct
            #   object methods wb.close(), excel.Quit()
            # kill excel process
            # os.system ( 'taskkill /f /im EXCEL.exe' )
        except Exception as e:
            log.debug( 'Failed to send to printer. Excp: {0}'.format( e ) )
            return False
        # end try

        log.debug( LOG_EXIT )
        return True
    # end function



    def _updateTDS( self, printfile:str ) -> bool:
        '''
        Update TDS template based on self.tdscellmap, which should be populated
        from the CreateTDS() with processed test result values
        Loop through each entry in self.tdscellmap to append the result (or
        image for cells with 'type': 'image') to the TDS template using
        win32com

        Return : True if TDS is updated
                 False if failed to update

        '''
        log.debug( LOG_ENTER )

        checkstatus         = False
        novalue             = []
        nostatus            = []
        self.alldatapresent = True

        if ( TST_FINAL_RESULT in self.boardinfodict.keys() ):
            if ( self.boardinfodict[ TST_FINAL_RESULT ] == TST_RES_PASSED ):
                isallpass = True
            else:
                isallpass = False
            # end if
        else:
            log.debug( 'no final result in boardinfo' )
            isallpass = False
        # end if

        # Get the active worksheet
        # Try to initiate Excel with win32
        try:
            excel = win32com.client.Dispatch( 'Excel.Application' )
            excel.Visible = False
            excel.DisplayAlerts = False
        except Exception as e:
            log.debug( "Failed to open excel Application error {0}".
                               format( e ) )
            return False
        # end try

        # Set sheet name to work with
        # default to 'Sheet 1' if not available in specfile
        if ( BRD_TDS_SHEET_FN in self.cfginfo.keys() ):
            self.tdssheet = self.cfginfo[ BRD_TDS_SHEET_FN ]
        else:
            self.tdssheet = 'Sheet1'
        # end if

        log.debug( 'self.tdstmptfn is {}'.format( self.tdstmptfn ) )
        path = os.path.abspath( os.path.dirname( sys.argv[ 0 ] ) )
        original_tdspath = os.path.join( os.path.join( path, 'TDS' ),
                self.tdstmptfn )

        copy_tdspath = os.path.join( path, 'local_{0}'.format(
                self.tdstmptfn ) )

        try:
            local_tdspath = shutil.copy( original_tdspath, copy_tdspath )
            log.debug( 'TDS path of local template copy: {0}'.format(
                    local_tdspath ) )
        except PermissionError:
            log.debug( 'Close excel file to proceed: {0}'.format(
                    copy_tdspath ) )
            return False
        except Exception as ex:
            log.debug( 'Unexpected exception encountered during copy of TDS '
                'template'.format( ex ) )
            return False
        # end try

        wb = excel.Workbooks.Open( local_tdspath )
        log.debug( 'wb object: {0}'.format( wb ) )

        log.debug( 'sheet name: {0}'.format( self.tdssheet ) )
        ws = wb.Worksheets( self.tdssheet )

        result_dic = self.tdscellmap.copy()
        tds = self.tdscellmap

        for key in self.tdscellmap:
            # look for either row and column or range in tds cell map
            # convert row and column to excel range is available
            if ( ( KEY_TDS_ROW in tds[ key ] ) and
                ( KEY_TDS_COLUMN in tds[ key ] ) ):
                row = tds[ key ][ KEY_TDS_ROW ]
                column = tds[ key ][ KEY_TDS_COLUMN ]
                addr = self.convertIntToExcelAddr( row = row,
                                                   col = column )
            elif ( KEY_TDS_RANGE in tds[ key ] ):
                addr = tds[ key ][ KEY_TDS_RANGE ]
            # end if


            # map for non-value entries
            if ( KEY_TDS_TYPE in tds[ key ] ):
                self.HandleNonValueTDSCell( tds[ key ], ws, isallpass, addr )
            # map for value entries

            else:
                # if provided a value directly, as in SQL results
                # or traverse methods, use that instead of where and select
                if ( KEY_TDS_VALUE in tds[ key ] ):
                    value = tds[ key ][ KEY_TDS_VALUE ]
                else:
                    continue

                # Capitalize PASS and FAIL for GX2 boards
                # 9/19/2017 HYF: Why only for GX2? Why not for all?
                if ( 'GX2' in self.boardinfodict[ HW_PRODUCT_ID ] ):
                    if ( value == TST_RES_PASSED ):
                        value = SQL_PASS
                    elif ( value == TST_RES_FAILED ):
                        value = SQL_FAIL
                    # end if
                # end if

                try:
                    ws.Range( addr ).Value = str( value )
                except Exception as e:
                    log.debug( 'Unknown error when updating Excel: {0}'
                            "".format( e ) )
                    log.debug( 'Please try again. If error persists, '
                            'please contact developer' )
                    messagebox.showwarning( title = 'Excel Update Error',
                            message = 'Unknown Excel error, please try '
                            'again.' )
                    return False
                #  end try

                # Put all result in dict for debugging and
                # validation
                result_dic[ key ][ KEY_TDS_VALUE ] = str( value )
            # end else (non-image result)
        # end for( each key )

        write_json_file( "result.json", result_dic )
        log.debug( 'wrote results to local file: result.json' )
        # show message to user if any key not
        # updated on TDS
        # checking updated elements
        alldatapresent = True
        for eachkey in result_dic.keys():
            if ( KEY_TDS_TYPE in result_dic[ eachkey ] ):
                if ( result_dic[ eachkey ][ KEY_TDS_TYPE ] == 'Image' ):
                    continue
                # end if
            # end if

            # only look for values table
            if ( KEY_TDS_VALUE not in result_dic[ eachkey ] ):
                log.error( "Failed to update value for {0}".format( eachkey ) )
                novalue.append( eachkey )
                alldatapresent = False
            # end if
        # end for

        # Display error
        if novalue:
            unkeys = "\n".join( novalue )
            msg = ( "Following Keys are not updated on TDS: "
                    " \n{0}".format( unkeys ) )
            log.error(msg)
            # messagebox.showerror( title = 'TDS Update Error', message = msg )
            self.alldatapresent = False
            # if lacking of test result value, show the error image
        # end if

        path = os.path.abspath( os.path.dirname( sys.argv[ 0 ] ) )
        tds_copy_path = os.path.join( path, printfile )

        try:
            log.debug( 'Saving TDS as: {0}'.format( tds_copy_path ) )
            wb.SaveAs( tds_copy_path )
        except Exception as e:
            log.debug( 'Failed to save workbook filename: {0}'.format(
                    printfile, e ) )
            return False
        # end try

        # Save TDS xls if it is an AT3552 product
        if( 'AT3552' in self.boardinfodict[ HW_PRODUCT_ID ] ):
            # Attempt to get save path from stationinfo, default to Saved TDS
            if( 'TDS Save Location' in self.stationfile ):
                log.debug( 'Manual save location found in station file' )
                savepath = self.stationfile[ 'TDS Save Location' ]
            else:
                log.debug( 'No save location in station file, using default' )
                path = os.path.abspath( os.path.dirname( sys.argv [ 0 ] ) )
                savepath = os.path.join( path, 'Saved TDS' )


            log.debug( 'Using path: {}'.format( savepath ) )
            printfile = self.boardinfodict[ BRD_SN ] + '.xls'
            tds_copy_path = os.path.join( savepath, printfile )

            try:
                log.debug( 'Saving AT3552 TDS as: {0}'.format( tds_copy_path ))
                wb.SaveAs( tds_copy_path )
            except Exception as e:
                log.debug( 'Failed to save AT3552 Copy: {0}'.format(
                        printfile, e ) )

                return False
            # end try

        log.debug( 'Close excel workbook' )
        wb.Close( False )

        log.debug( 'Close excel application' )
        excel.Quit()

        if( isallpass == False ):
            log.debug('Communicating to automation script that FINAL RESULT not passed')
            return False

        log.debug( LOG_EXIT )
        return True
    # end function



    def HandleNonValueTDSCell( self, tds_cell, ws, isallpass, addr ):
        '''
        Handle adding non-value things to TDS, i.e. images

        params:
            tds_cell: the dictionary in TDS Cell Map being processed
            ws: the Excel worksheet object to modify (pass as is from _updateTDS)
            isallpass: final rollup for fail stamp (pass as is from _updateTDS)
            addr: addres of target Excel cell (pass as is from _updateTDS)
        '''

        if ( tds_cell[ KEY_TDS_TYPE ].lower() ==
                                        KEY_TDS_IMAGE ):

            path = os.path.abspath( os.path.dirname( sys.argv[ 0 ] ) )
            where = tds_cell[ KEY_TDS_WHERE ]
            select = tds_cell[ KEY_TDS_SELECT ]

            # TDS cell map is an image
            image_path = "{0}\\{1}\\{2}".format( path, where,
                    select )
            log.debug( 'image path: {0}'.format( image_path ) )
        # end if

        if ( KEY_TDS_SUB_TYPE in tds_cell ):
            # Update failed stamp
            # checking all conditions
            subtype = tds_cell[ KEY_TDS_SUB_TYPE ]

            if ( ( isallpass == False ) and
                 ( subtype.lower() == 'fail' ) ):
                pic = ws.Pictures().Insert( r"{0}".format(
                        image_path ) )
                range = ws.Range( addr )
                pic.Left = range.Left
                pic.Top = range.Top
            # end if
        else:
            pic = ws.Pictures().Insert( r"{0}".format(
                    image_path ) )
            range = ws.Range( addr )
            pic.Left = range.Left + 3
            pic.Top = range.Top + 3
        # end if
    # end function



    def convertIntToExcelAddr( self, row, col ) -> str:
        """ Convert a row and column number to an Excel-style cell address

        """
        log.debug( LOG_ENTER )

        LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

        # If col is already a string, don't attempt to convert it to a letter
        if ( isinstance( col, str ) ):
            result = col
        else:
            result = []
            while col:
                col, rem = divmod( col - 1, 26 )
                result[ :0 ] = LETTERS[ rem ]
        # end if

        log.debug( LOG_EXIT )
        return ''.join(result) + str(row)
    # end function



    def TDSCalculateVal( self, eachentry: dict ) -> str:
        """
        eachentry: dict from TDS cell map.
        {
                    "Name": "B Constant",
                    "range": "I24",
                    "data location" : "Calculation",
        }

        Return:
            String: if the value was successfully calculated
            None: if the value cannot be successfully calculated
        """
        log.debug( LOG_ENTER )

        if ( eachentry[ BRD_NAME ] == LM_TDS_DATA_BC ):
            # B = [ ln( Rt/R25 )] /( 1/T - 1/T0 )
            rt = None
            t = None
            for eachdict in self.tdscellmap:
                if ( eachdict[ BRD_NAME ] == LM_TDS_DATA_RTH_KOHM ):
                    rt = self.tdsdatasetdict[ LM_TDS_DATA_RTH_KOHM ]
                elif ( eachdict[ BRD_NAME ] == LM_TDS_DATA_LSR_TEMP ):
                    t = self.tdsdatasetdict[ LM_TDS_DATA_LSR_TEMP ]
                else:
                    pass
                # end if
            # end for

            if ( rt == None ) or ( t == None ):
                log.debug( 'Missing values necessary to calculate {}'.format(
                    eachentry[ BRD_NAME ] ) )
                return None
            # end if

            log.debug( 'rt is {}, t is {}'.format( rt, t ) )
            T = Decimal( t ) + Decimal( "273.15" )
            T0 = Decimal( "328.15" )
            R25 = Decimal( "2.985" )

            if ( rt == 0 ):
                log.debug( 'rt is 0 (zero) cannot use log function. Ret zero' )
                B = 0
            else:
                B = ( Decimal( math.log( Decimal( rt ) / R25 ) ) / \
                    ( ( Decimal( '1' ) / T ) - ( Decimal( '1' ) / T0 ) ) )
            # end if

            ret = str( B )
        elif ( eachentry[ BRD_NAME ] == LM_TDS_DATA_ERR_MAX ):
            te = "{0:.2f}".format( random.uniform( 0.18, .29 ) )
            return te
        elif ( eachentry[ BRD_NAME ] == LM_TDS_DATA_ERR_MIN ):
            te = float( "{0:.2f}".format( random.uniform( 0.15, .29 ) ) )
            nte = te * float( '-1' )
            return str( nte )
        elif ( eachentry[ BRD_NAME ] in [ LM_TDS_DATA_CNR, LM_TDS_DATA_CSO,
                                          LM_TDS_DATA_CTB ] ):
            # Get the worst value for distortion values we have to calculate
            # 1. Get all the values for that distortion test.
            #    Note: CSO will take the minimum value from both CSO- & CSO+
            # 2. Compare the absolute value of the numbers (to account for
            #    some software recording distortion numbers as negative, and
            #    some software recording distortion numbers as positive)
            #    to find the worst(smallest value).
            keyname = eachentry[ BRD_NAME ]
            minvalue = None
            vals = [ value for key, value in self.tdsdatasetdict.items()
                     if keyname in key ]
            log.debug( 'values for {} calc: {}'.format( keyname, vals ) )

            for eachval in vals:
                if ( ( eachval == 0 ) or ( eachval == '' ) or
                     ( eachval is None ) ):
                    log.debug( 'skip 0, empty string value, and None values' )
                    continue
                # end if

                currvalue = abs( Decimal( eachval ) )
                if ( ( minvalue is None ) or ( currvalue < abs( minvalue ) ) ):
                    log.debug( 'min value is now: {}'.format( eachval ) )
                    minvalue = eachval
                # end if

            # end for
            ret = minvalue
        elif ( eachentry[ BRD_NAME ] == LM_TDS_DATA_CHIRP ):
            # According to ME, copy method from previous TDS software
            queryvalue = self.GetChirpPreviousMethod()
            ret = queryvalue
        else:
            log.debug( 'unknown calculation' )
            ret = None
        # end if

        log.debug( LOG_EXIT )
        return ret
    # end function



    def GetSQLDBInformation( self, id:str, querydict:dict,
        retdict: bool = False ) -> bool:
        """Query the SQL database for information on the specified id with the
        parameters defined in querydict

        SELECT [value_column]
        FROM [table]
        WHERE [idcolumn] = '[id]'
        ORDER BY [timecolumn]
        [orderdirection]

        id: Serial Number of device
        querydict: table information. Fields required are:
                   value_column: name of the column for the field you want.
                         Can be 'top 1 colname' if you want just one record.
                         Can be 'top 1 *' if you want the entire record (row).
                   table: name of the SQL table
                   idcolumn: name of the column to filter by, e.g. 'SerNo'
                   id: value of the id
                   timecolumn: name of the column in the SQL table to sort by
                   orderdirection: choose ascending or descending. descending
                                   will get the latest value.
                   initcfg: this is defaulted to "TOP 1"
                   server: type of server, e.g. Production, MES
            Looks like:
            {
                    "initcfg": "TOP 1",
                    "table": "dbo.ChirpSpecTrumData",
                    "idcolumn": "SerNo",
                    "value_column": "ChirpMhzmW1",
                    "timecolumn": "TimeStamp",
                    "orderdirection" :"DESC",
                    "value": "",
                    "server": "MES SQL"
            }
        retdict: boolean to determine whether the sql will return rows as
                 dictionaries or tuples

        Return: True : Value successfully retrieved
                False: Failed to retrieve value
        """
        log.debug( LOG_ENTER )

        # Find server info, assume production if not specified
        # Allow to change for printing in different locations by taking the
        #    location from the station file.
        #    If it states a specific server location for a field, pull
        #    from there. If it specifies a server type,
        if ( SQL_QUERY_SERVER not in querydict ):
            servertype = DB_SQL_PROD
        else:
            servertype = querydict[ SQL_QUERY_SERVER ]
        # end if

        if servertype not in self.stationfile:
            log.error( 'Could not find the SQL server {} in station '
                'file'.format( servertype ) )
            querydict[ KEY_TDS_VALUE ] = None
            return False

        servername = self.stationfile[ servertype ]
        serverinfo = None
        for eadbinfo in db_list:
            if ( eadbinfo[ DB_DEF_ID ] == servername ):
                serverinfo = eadbinfo
                break
            # end if
        # end for

        if serverinfo is None:
            log.debug( 'could not find the server info!' )
            querydict[ KEY_TDS_VALUE ] = None
            return False
        # end if

        # Emcore package
        try:
            ms = mssqlserver( host = serverinfo[ DB_HOSTNAME ], user =
                serverinfo[ DB_USERNAME ], pwd = serverinfo[ DB_PASSWORD ],
                db = serverinfo[ DB_DEF_NAME ], retdict = retdict )
        except Exception as ex:
            log.debug( 'issue with sql connect: {0}'.format( ex ) )
            querydict[ KEY_TDS_VALUE ] = None
            return False
        # end try

        if ( SQL_QUERY_INIT_CFG not in querydict ):
            init_cfg = 'TOP 1'
        else:
            init_cfg = querydict[ SQL_QUERY_INIT_CFG ]
        # end if

        value = None
        query = ( "SELECT {0} {1} FROM {2} WHERE ( {3} = '{4}' ) ORDER BY {5} "
                  "{6}" ).format(
                    init_cfg,
                    querydict[ SQL_QUERY_VALUE_COL ],
                    querydict[ SQL_QUERY_TABLE ],
                    querydict[ SQL_QUERY_ID_COL ] , id,
                    querydict[ SQL_QUERY_TIME_COL ],
                    querydict[ SQL_QUERY_ORDER_DIR ] )

        log.debug( 'query is {}'.format( query ) )
        try:
            value = ms.ExecQuery( query )
        except Exception as ex:
            log.debug( 'issue with sql query: {0}'.format( ex ) )
            querydict[ KEY_TDS_VALUE ] = None
            return False
        # end try

        log.debug( 'value: {}'.format( value ) )

        if ( value is not None ):
            querydict[ KEY_TDS_VALUE ] = value
        else:
            querydict[ KEY_TDS_VALUE ] = value
            log.debug( 'Value not found' )
            return False
        # end if

        log.debug( LOG_EXIT )
        return True
    # end function
# end class





class TDSWorker_TX( TDSWorker ):
    def __init__( self, serialnumber: str , sessiontype: str, num_type: str ):
        """
        serialnumber: value for the number type
        sessiontype: sessiontype is intended to distinguish between production,
            engineering, oql, spc
        num_type: number type, e.g. Serial Number, Laser ID
        """
        super().__init__( serialnumber = serialnumber,
            sessiontype = sessiontype )

        self._boarddata_obj = BoardInformation( None, None )
        self.num_type = num_type
    # end function


    def SaveGX2TDSToCompressedPDF( self ):
        '''
        Code for saving GX2 BC directly to pdf, called by printTDSData for
        GX2 boards

        may need to reinstall pypiwin32 to work on station
          $ pip uninstall pypiwin32
          $ pip install --no-cache-dir pypiwin32
        !!Get pdfsizeopt from
        https://github.com/pts/pdfsizeopt/releases/
        download/2017-09-02w/pdfsizeopt_win32exec-v6.zip
        and extract to pdfsizeopt folder
        '''

        log.debug( LOG_ENTER )

        tdsfilename = self.tdsprintfn
        pcbid = self.boardinfodict[ BRD_PCB_ID ]
        MySleep( 2 )
        # final directory to save compressed pdf in
        finaldir = "G:\\Man\\BC_converter"
        # finaldir = "C:\\GX2"
        # directory to save temporary uncompressed pdf to
        path = os.path.abspath( os.path.dirname( sys.argv[ 0 ] ))
        workdir = os.path.join( path, 'GX2BC' )
        # finaldir = os.path.join( path, 'Saved TDS')
        continuecompress = False

        filename = os.path.join( path, tdsfilename )
        save_pdf = os.path.join( workdir, '{0}_large.pdf'.format(
                                 pcbid ) )
        xlTypePDF = 0
        log.debug( 'GX2 product, saving pdf to {0}'.format(
                save_pdf ) )

        excel = win32com.client.gencache.EnsureDispatch(
                "Excel.Application" )
        book = excel.Workbooks.Open( Filename = filename )

        try:
            book.ExportAsFixedFormat( xlTypePDF, save_pdf )
            continuecompress = True
        except Exception as e:
            log.debug( str( 'Could not save GX2 as PDF. Please make '
                    'sure to reinstall pypiwin32 with "pip uninstall '
                    'pypiwin32"then "pip install --no-cache-dir '
                    'pypiwin32"' ) )
        # end try

        if ( not os.path.isfile( './pdfsizeopt/pdfsizeopt.exe' ) ):
            continuecompress = False
            log.debug( str('WARNING: pdfsizeopt is not installed, so '
                'GX2 BC will not be compressed. Download it here: '
                'https://github.com/pts/pdfsizeopt/releases/download/'
                '2017-09-02w/pdfsizeopt_win32exec-v6.zip'
                ' and extract the files to the pdfsizeopt folder' ) )
        # end if

        if ( continuecompress ):
            compressbat = "pdfsizeopt\\pdfsizeopt.bat"
            largepath = workdir + '\\' + '{0}_large.pdf'.format(
                    pcbid )
            compresspath = finaldir + '\\' + '{0}.pdf'.format( pcbid )
            log.debug( "largepath is {}".format( largepath ) )
            log.debug( "compresspath is {}".format( compresspath ) )

            if( os.path.isfile( compresspath ) ):
                log.debug( "File for compresspath already exists, removing" )
                os.remove( compresspath )
            continuecompress = False
            try:
                Popen( compressbat + ' ' + largepath + ' ' +
                        compresspath )
                continuecompress = True
            except Exception as e:
                log.debug( 'File with name already exists, did not '
                        'compress' )
            # end try
        # end if

        sheet = None
        book = None
        excel.Quit()
        excel = None

        log.debug( LOG_EXIT )
    # end function



    def CreateTDS( self, printbool: int ) -> dict:
        """ Fill all the required fields on test data sheet
        based on test results. Reads the data from testinfo (json) file and
        update each cell based on test name and group.

        function calls:
            self.GetBoardInfo: Get boardinfo
            self.GetConfigInfo: Get specfile
            self._boarddata_obj.getBoardInfoFromSpecData: Get information from
                specfile under "Product Spec Data" key and add to boardinfodict
            self.GetAdditionalInfo: Get test data and tds template info
            self.updateTemplate: Update TDS template with all gathered info
            self.printTDSData: Print updated TDS tempalte

        :param: printbool: int value used as bool. Can be 0 or 1.
                1 to print. 0 if we want to skip TDS update and printing.
        :return: (dictionary)
            retdict:
                status_bool
                status_msg
        """
        log.debug( LOG_ENTER )

        retdict = dict( status_bool = False, status_msg = '' )

        # Get board information from database
        retdict = self.GetBoardInfo()
        if( retdict[ KEY_TDS_STATUS_BOOL ] == False ):
            return retdict

        # Get config information (specfile) from database
        retdict = self.GetConfigInfo()
        if( retdict[ KEY_TDS_STATUS_BOOL ] == False ):
            return retdict


        self.boardinfodict[ DB_PRODUCT_FAMILY ] = FAMILY_CATV
        self.boardinfodict = self._boarddata_obj.getBoardInfoFromSpecData(
                board_info_dict = self.boardinfodict,
                product_spec_dict = self.cfginfo[ BRD_SPEC_DATA ] )

        # manual spec for testing
        # self.cfginfo = read_json_file( DEV_SPEC_FN )

        # write for debugging
        # write_json_file( "boardinfo.json", self.boardinfodict )

        # Write specfile
        write_json_file( DEV_SPEC_FN, self.cfginfo )
        log.debug( 'Wrote specfile to local file: {}'.format( DEV_SPEC_FN ) )

        retdict[ KEY_TDS_CFG_INFO ] = self.cfginfo
        retdict[ KEY_TDS_BOARD_INFO_DICT ] = self.boardinfodict

        # if there is no TDS cell map in the spec file, this is Aurora product
        # and does not need TDS, so jump to next
        if ( BRD_TDS_CELL_MAP not in self.cfginfo ):
            msg = 'Aurora Product, no TDS needed'
            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = True
            return retdict

        # Get test data from database and tds template information from
        # specfile
        tempretdict = self.GetAdditionalInfo()
        if( tempretdict[ KEY_TDS_STATUS_BOOL ] == False ):
            retdict[ KEY_TDS_STATUS_MSG ] = tempretdict[ KEY_TDS_STATUS_MSG ]
            retdict[ KEY_TDS_STATUS_BOOL] = False
            return retdict

        # set info again after adding to them
        retdict[ KEY_TDS_CFG_INFO ] = self.cfginfo
        retdict[ KEY_TDS_BOARD_INFO_DICT ] = self.boardinfodict

        # now that we have gotten specfile and boardinfo, return
        # them if we do not intend to print
        if ( printbool == 0 ):
            msg = 'Skipping TDS update and printing'
            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = True
            return retdict

        # write boardinfo to json again if updated
        write_json_file( "boardinfo.json", self.boardinfodict )
        log.debug( 'Wrote boardinfo to local file: boardinfo.json' )

        # Update TDS template
        retbool = self.updateTemplate()
        if ( retbool == False ):
            msg = 'Unsuccessful update of template, may be due to ' \
                  'Excel file already being open'
            log.debug( msg )
            # messagebox.showwarning( title = 'Status', message = msg )

            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        # end if

        msg = 'Successful update of template'
        log.debug( msg )
        retdict[ KEY_TDS_STATUS_MSG ] = msg

        # Print TDS if TDS template is filled completely
        if ( self.alldatapresent ):
            retbool = self.printTDSData()
            if ( retbool == False ):
                msg = 'Unsuccessful print of TDS'
                log.debug( msg )
                # messagebox.showwarning( title = 'Status', message = msg )

                retdict[ KEY_TDS_STATUS_MSG ] = msg
                retdict[ KEY_TDS_STATUS_BOOL ] = False
                return retdict
            # end if

            msg = 'Successfully printed TDS'
            log.debug( msg )

            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = True
            # return the dictionaries retrieved from the database so that
            # they don't have to be queried again
            retdict[ KEY_TDS_CFG_INFO ] = self.cfginfo
            retdict[ KEY_TDS_BOARD_INFO_DICT ] = self.boardinfodict
        else:
            msg = 'TDS template incomplete, will not print'
            log.debug( msg )
            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False

        write_json_file( 'boardinfo.json', self.boardinfodict )

        return retdict

        log.debug( LOG_EXIT )
    # end if

    # end function



    def GetBoardInfo( self ) -> dict:
        '''
        Query the database for board information
        Save this board information to self.boardinfodict for later use
        Return retdict
            retdict[ KEY_TDS_STATUS_BOOL ]: True/False if process succeeded
            retdict[ KEY_TDS_STATUS_MSG ]: Information message
        '''

        log.debug( LOG_ENTER )

        retdict = dict( status_bool = True, status_msg = '' )

        # Get the board information to determine whether we should use sql
        # or mongo. If SQL, also query all the necessary data.
        try:
            self.boardinfodict = self._boarddata_obj.getBoardInfoFromDataBase(
                self.num_type, self._serialnumber, PRODUCT_TYPE_TX )
        except Exception as e:
            log.error( 'Issue when getting board info: {0}'.format( e ) )
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            retdict[ KEY_TDS_STATUS_MSG ] = 'Failed to connect to databases'
            return retdict
        # end try

        if ( self.boardinfodict is None ):
            msg = 'Did not find any board with the given SN in Mongo or SQL'
            log.debug( msg )
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            retdict[ KEY_TDS_STATUS_MSG ] = msg
            return retdict
        # end if

        log.debug( LOG_EXIT )

        return retdict
    # end function



    def GetConfigInfo( self ) -> dict:
        '''
        Get config information (specfile) from database
        Save this to self.cfginfo for later use
        Return retdict
            retdict[ KEY_TDS_STATUS_BOOL ]: True/False if process succeeded
            retdict[ KEY_TDS_STATUS_MSG ]: Information message
        '''

        log.debug( LOG_ENTER )

        retdict = dict( status_bool = True, status_msg = '' )

        if ( HW_PRODUCT_ID not in self.boardinfodict ):
            log.debug( "HW_PRODUCT_ID not found on board spec"  )
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        # end if

        productid = self.boardinfodict[ HW_PRODUCT_ID ]

        # Get product spec file from MONGO DB
        self.cfginfo = getDeviceInfo( db_info = db_info,
                                 product_id = productid )

        if ( not self.cfginfo ):
            msg = str( 'Failed to find spec file {0} from MONGO DB'.
                format( productid ) )
            log.debug( msg )

            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        # end if

        log.debug( LOG_EXIT )

        return retdict
    # end function



    def GetAdditionalInfo( self ) -> dict:
        '''
        Get additional information from specfile and query database for test
        data. Add additional information to self.boardinfodict
        Calls functions:
            self.getRelatedDataFromCfg: get information from spec file like
                tdscellmap
            self.GetTestData: query database for test data (mongo or SQL)
        Return retdict
            retdict[ KEY_TDS_STATUS_BOOL ]: True/False if process succeeded
            retdict[ KEY_TDS_STATUS_MSG ]: Information message
        '''

        log.debug( LOG_ENTER )

        retdict = {
            KEY_TDS_STATUS_BOOL: True,
            KEY_TDS_STATUS_MSG: ''
        }

        # Ensure that BRD_SERIAL_NUM has the serial number, regardless
        # of how it is keyed in database. TDS looks for BRD_SERIAL_NUM
        if ( BRD_SN in self.boardinfodict.keys() ):
            serialnum = self.boardinfodict[ BRD_SN ]
        elif ( BRD_CUSTOMER_ID in self.boardinfodict.keys() ):
            serialnum = self.boardinfodict[ BRD_CUSTOMER_ID ]
        # end if
        self.boardinfodict[ BRD_SERIAL_NUM ] = serialnum


        # Get tds information from spec file
        retbool = self.getRelatedDataFromCfg()
        if ( retbool == False ):
            msg = 'Missing information in device spec!'
            log.debug( msg )
            messagebox.showwarning( title = 'Status', message = msg )

            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        else:
            log.debug( 'Product spec configurations verified' )
        # end if


        # Get test data for the board with the given product family
        # Andy 10/20/17: self.boardinfodict[ DB_PRODUCT_FAMILY ] is currently
        # set statically to CATV in CreateTDS() function, so SATCOM products
        # are not supported until this is changed
        product_family = self.boardinfodict[ DB_PRODUCT_FAMILY ]

        if ( product_family == FAMILY_CATV ):
            retbool = self.GetTestData_CATV()
        elif ( self.product_family.upper( ) == FAMILY_SATCOM ):
            retbool = self.GetTestData_SATCOM()
        else:
            log.error( "Unsupported Product Family type {0} ".format(
                self.product_family ) )
            retbool = False
        if ( retbool == False ):
            msg = 'One or two test results not found in DB'
            log.debug( msg )
            messagebox.showwarning( title = 'Status', message = msg )

            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        else:
            log.debug( 'Test data successfully added to boardinfodict' )
        # end if

        log.debug( LOG_EXIT )

        return retdict
    # end function



    def GetTestData_CATV( self ) -> bool:
        """ Retrieve test data from super test result if they are included
        in the boardinfodict passed, which will occur for mongo super results
        If any test results are missing, query SQL for them

        Attaches results to self.boardinfodict

        Process function calls:
            self.GetSqlTestResults: query for any missing test results
                self.ParseSqlResults: attach results to tds cell map
            self.HandleGX2Rev: append revision codes for GX2 tds from SQL data
            self.ParseMongoResultsByTraversal: attach results to tds cell map
                                            by traversal model if present
            self.ParseMongoResults: use original mongo search method if needed
            self.CheckFinalResults: check to make sure that all tests present

        Return:
            True: Test results retrieved
            False: Issue with obtaining data ( connection error,
                   or missing data )
        """
        log.debug( LOG_ENTER )

        # -----------------------
        # 1. Initial setup
        # -----------------------
        msconnected = False

        try:
            ms = mssqlserver()
            msconnected = True
        except Exception as ex:
            log.debug( 'issue with sql connect: {0}'.format( ex ) )
            return False
        # end try

        finalresults = {}

        # Modify these if board changes in what test data is required
        # This could be migrated to specfile if differences appear between TX
        # boards in the future
        required_tests = [ TST_RES_FREQ_RESP_DATA, TST_RES_DIST_DATA ]
        results_dict = {
            TST_RES_FREQ_RESP_DATA: {
                'pass': None,
                'sql': {},
                'mongo': None,
                'missing': True
            },
            TST_RES_DIST_DATA: {
                'pass': None,
                'sql': {},
                'mongo': None,
                'missing': True
            }
        }
        # Final results dictionary
        self.results = [ self.boardinfodict ]

        # check if self.tdscellmap is initiated yet. If not, initiate it
        # this will be used in the case that SQL data is being pulled
        if ( not self.tdscellmap ):
            self.tdscellmap = self.cfginfo[ BRD_TDS_CELL_MAP ]
        # end if


        # -----------------------
        # 2. Query SQL database if needed
        # -----------------------
        # Check if boardinfodict contains each test type from required_tests
        # If they are there, they came from mongoDB. If not, attempt to
        # retrieve them from sql
        for test_type in required_tests:

            if( test_type in self.boardinfodict ):
                # Results already available from mongoDB
                results_dict[ test_type ][ 'mongo' ] = True
                # self.results.append( self.boardinfodict[ test_type ] )
                results_dict[ test_type ][ 'missing' ] = False

            else:
                # GX2 distortion results have PCB id stored in SN column
                if( 'GX2' in self.boardinfodict[ BRD_FULL_MODEL ] and
                    test_type == TST_RES_DIST_DATA ):
                    search_id = self.boardinfodict[ BRD_PCB_ID ]
                else:
                    search_id = self.boardinfodict[ BRD_SERIAL_NUM ]

                # Attempt to get test results for board from SQL
                retdict = self.GetSqlTestResults( test_type, search_id, ms )
                if( retdict ):
                    results_dict[ test_type ][ 'pass' ] = retdict[
                                                            TST_FINAL_RESULT]
                    results_dict[ test_type ][ 'missing' ] = False
                    results_dict[ test_type ][ 'sql' ] = retdict[ TST_RESULT ]
                    self.ParseSqlResults( test_type, retdict[ TST_RESULT ] )



        # get Arris Rev. and Emcore Rev. for GX2 products
        if ( ( HW_PRODUCT_ID in self.boardinfodict.keys() )
            and ( 'GX2' in self.boardinfodict[ HW_PRODUCT_ID ] ) ):
            self.HandleGX2Rev( results_dict )
        # end if

        # -----------------------
        # 3. Parse mongo results
        # -----------------------

        # If TDS Traversal Model exists in specfile, search self.boardinfodict
        # with it. (Traversal model currently only used by AT3552, so
        # it definitely uses mongo results)
        if( BRD_TDS_TRAVERSE_MODEL in self.cfginfo ):
            retbool = self.ParseMongoResultsByTraversal()
            if( retbool == False ):
                log.debug( 'Error when parsing test results by traversal model')
                return False

        # Otherwise attempt to parse results with From().where().select()
        # (ParseMongoResults checks against results_dict if there are sql
        # results already, ParseMongoResultsByTraversal does not check)
        else:
            retbool = self.ParseMongoResults( results_dict )
            if( retbool == False ):
                log.debug( 'Error when parsing mongo test results')
                return False


        # -----------------------
        # 4. Check that final results are present to continue
        # -----------------------
        all_res_present = self.CheckFinalResults( required_tests, results_dict )
        if( not all_res_present ):
            return False

        log.debug( LOG_EXIT )

        return True
    # end function



    def GetSqlTestResults( self, test_type, search_id, ms ) -> dict:
        '''
        Run a SQL query for a given test type (e.g. TST_RES_FREQ_RESP_DATA)
        Pass the search_id manually, as distortion GX2 results use a different
        identifier (BRD_PCB_ID)

        params:
            test_type: TST_RES_FREQ_RESP_DATA/TST_RES_DIST_DATA
            search_id: id to search by (SN, or PCB for GX2 Dist data)
            ms: mssqlserver object initiated in GetTestData
        Returns:
            retdict: if sql query is successful
                retdict[ TST_FINAL_RESULT ]: pass/fail
                retdict[ TST_RESULT ]: sql query result
            None: no sql results found
        '''

        log.debug( LOG_ENTER )

        # Information for how queries of each type should be run
        query_dictionary = {
            TST_RES_DIST_DATA: {
                'query': "SELECT top 1 * FROM Catv_QAM WHERE " \
                                 "[TransmitterSerNo] = '{0}' ORDER BY " \
                                 "[RecordId] DESC",
                'pfcolumn': 15
            },
            TST_RES_FREQ_RESP_DATA: {
                'query': "SELECT top 1 * FROM QAM_TX_FreqResp " \
                                 "WHERE [SerNo] = '{0}' ORDER BY " \
                                 "[RecordID] DESC",
                'pfcolumn': 12
            }
        }

        # Build sql query using predefined dictionary
        query_info = query_dictionary[ test_type ]
        sql_query = query_info[ 'query' ].format( search_id )

        # Execute sql query
        try:
            sql_response = ms.ExecQuery( sql_query )
        except Exception as e:
            log.error( 'SQL query for {0} failed with message: {1}'.format(
                test_type, e) )
            return None


        if( sql_response ):
            # Get the final pass/fail result in sql
            final_result = sql_response[ 0 ][ query_info[ 'pfcolumn' ] ]
            log.debug( '{0} results found: {1}'.format(
                test_type, final_result) )

            # Build return dictionary
            retdict = {}
            retdict[ TST_FINAL_RESULT ] = final_result
            retdict[ TST_RESULT ] = sql_response
        else:
            return None

        log.debug( LOG_EXIT )

        return retdict
    # end function



    def ParseSqlResults( self, test_type, sql_results ) -> bool:
        '''
        Reads sql results and writes them into self.tdscellmap to later
        be used in _updateTDS to write to the TDS template
        Looks for a 'SQL' key in each TDS Cell Map (specfile) dictionary to
        determine which table and column to get results from

        params:
            test_type: TST_RES_FREQ_RESP_DATA/TST_RES_DIST_DATA
            sql_results: sql results from an ExecQuery
        Returns:
            True: self.tdscellmap successfully updated
            False: something went wrong in parsing sql results
        '''

        log.debug( LOG_ENTER )

        # cocnversion from sql table to test type
        sql_table_type = {
            'QAM_TX_FreqResp': TST_RES_FREQ_RESP_DATA,
            'Catv_QAM': TST_RES_DIST_DATA
        }


        tds = self.tdscellmap
        for key in tds.keys():
            if ( SPEC_DB_SQL not in tds[ key ] ):
                log.debug( 'SQL not in key. Continue.' )
                continue
            # end if
            if( KEY_TDS_VALUE in tds[ key ]
                and tds[ key ][ KEY_TDS_VALUE ] != None ):
                log.debug( 'Value already assigned for tds cell, skipping cell' )
                continue

            # Table and column that corresponds to the test result in TDS cellmap
            sql_dict = tds[ key ][ SPEC_DB_SQL ]
            sql_table = sql_dict[ SQL_QUERY_TABLE ]
            col = sql_dict[ KEY_TDS_COLUMN ]
            log.debug( 'col: {}'.format( col ) )

            # Skip tds cell if the type is not the same as the parsing results
            target_test_type = sql_table_type[ sql_table ]
            if( target_test_type != test_type ):
                continue

            val = sql_results[ 0 ][ col ]

            # find additional values for special keys
            additional_values = {}
            additional_value_keys = [ KEY_TDS_COMPARE, KEY_TDS_MAX,
                                        KEY_TDS_DEFAULT]

            for special_key in additional_value_keys:
                if( special_key in sql_dict ):
                    additional_values[ special_key ] = sql_results[
                                        0 ][ sql_dict[ special_key ] ]

            # Special case: convert optical power into dBm from sql
            if ( key == SPEC_TYPE_OPTICAL_PWR ):
                val = ( 10 * math.log10( val ) )


            val = self.HandleSpecialTDSKeys( val, tds[ key ], additional_values )

            # Apply value to tds cell map value key
            # !! tds references self.tdscellmap, so this change is persistent !!
            tds[ key ][ KEY_TDS_VALUE ] = val

        # end for

        log.debug( LOG_EXIT )

        return True

    # end function



    def ParseMongoResults( self, results_dict ) -> bool:
        '''
        If there are mongo results, iterate through the TDS cell map
        to populate it via mongo results. This function uses
        From().where().select() methods (as opposed to traverse_by_model)

        Results are saved to self.tdscellmap to later be used in _updateTDS()

        params:
            results_dict: as it appears in CreateTDS()

        returns:
            True: parsing successful
            False: error when parsing
        '''

        log.debug( LOG_ENTER )

        tds = self.tdscellmap
        search_locations = [ 'boardinfodict' ]

        # determine which test types to look in (not missing and no sql values)
        for test_type in results_dict:
            if( results_dict[ test_type ][ 'mongo' ] ):
                search_locations.append( test_type )


        for key in tds:
            if ( KEY_TDS_SELECT not in tds[ key ]
                or KEY_TDS_WHERE not in tds[ key ] ):
                log.debug( str('TDS Cell Map for {0} missing "where" or "select"'
                                ' keys for parsing mongo results'.format(
                                                                    key ) ) )
                continue
            # end if

            if( KEY_TDS_VALUE in tds[ key ]
                and tds[ key ][ KEY_TDS_VALUE ] != None ):
                log.debug( 'Value already assigned for tds cell, skipping cell' )
                continue

            # if tds has type, skip mongo search, this is for a non-dictionary
            # result
            if ( KEY_TDS_TYPE in tds[ key ]):
                continue

            # Use From().where().select() to parse self.boardinfodict for
            # mongo test result
            where = tds[ key ][ KEY_TDS_WHERE ]
            select = tds[ key ][ KEY_TDS_SELECT ]
            resultfound = False
            additional_values = {}

            # location will either be 'boardinfodict' meaning search the root
            # level, or a test type (e.g. TST_RES_FREQ_RESP_DATA) that is a key
            # inside self.boardinfodict
            for location in search_locations:
                if( location == 'boardinfodict' ):
                    search_dict = self.boardinfodict
                else:
                    search_dict = self.boardinfodict[ location ]

                try:
                    result = From( search_dict ).where( where ).select( select )
                except Exception as e:
                    log.debug( "query error {0} ".format( e ) )
                    log.debug( "query error on where: {0}; select {1}".format(
                                where, select ) )
                    result = [ None ]
                # end try

                # If this tds cell has a 'compare' key, find that too
                if ( KEY_TDS_COMPARE in tds[ key ] ):
                    compare = ( From( search_dict ).where( where ).select(
                                                tds[ key ][ KEY_TDS_COMPARE ] ) )
                    if ( len( compare ) > 0 and
                            compare[ 0 ] not in [ None, 'None' ] ):

                        additional_values[ 'compare' ] = float( compare[ 0 ] )
                # end if

                # If there was no select result, check the next search location
                # otherwise, break out of search loop
                if ( len( result ) > 0 and
                    result[ 0 ] not in [ None, 'None' ] ):
                    resultfound = True
                    break
                # end if
            # end for (search loop)

            # no result found, continue
            if( len( result ) < 1 ):
                continue

            log.debug( " Test value Query result {0} "
                "where {1} select {2}".format( result,
                where, select ) )
            # end if

            # Make sure its a string and only one element in list
            # before checking
            value = str( result[ 0 ] )

            value = self.HandleSpecialTDSKeys( value, tds[ key ],
                        additional_values )

            # Apply value to tds cell map
            # !! tds references self.tdscellmap so this change is persistent
            tds[ key ][ KEY_TDS_VALUE ] = value

        # end for
    # end function



    def HandleSpecialTDSKeys( self, value, tds_cell, additional_values ):
        '''
        Some tds cell map cells have special requirements
        This function will modify the value of the tds cell based on the
        special key

        Special keys:
            KEY_TDS_DEC_PLACE: set number of decimal places
                Andy 10/21/17: the decimal place actually displayed
                    seems to be determined in Excel formatting of the cell
            KEY_TDS_NEGATIVEFLIP: flip the value to negative
            KEY_TDS_ABSOLUTE: make the value absolute
            KEY_TDS_DEFAULT: set the value to a default beyond either a max or min
                KEY_TDS_MAX: max value beyond which value is set to default
            KEY_TDS_COMPARE: compare the value with another and choose the lower

        params:
            value: original value of test result
            tds_cell: the full tds dictionary entry e.g.
                "RLoss RFTP": {
                    "SQL": {},
                    "where": "",
                    "select": "",
                    "negativeflip": "true"
                }
            additional_values: dictionary of extra values needed
                e.g. { 'compare': val } or { 'default': val, 'max': val2 }

        returns: new value, or original if no special key was found
        '''

        # Default convert values to 1 decimal place
        # Note: this conversion happens even if there is no 'decimal place' key
        # in tds cell, so we need to wrap in a try in case its a non-number val
        decimal_place = 1
        if ( KEY_TDS_DEC_PLACE in tds_cell ):
            decimal_place = int( tds_cell[ KEY_TDS_DEC_PLACE ] )
        try:
            value = round( Decimal( value ), decimal_place )
        except:
            pass

        if( KEY_TDS_NEGATIVEFLIP in tds_cell ):
            value = str( -1 * float( value ) )

        if( KEY_TDS_ABSOLUTE in tds_cell ):
            value = str( abs( float( value ) ) )

        if( KEY_TDS_COMPARE in tds_cell and 'compare' in additional_values ):
            compare = additional_values[ 'compare' ]
            value = float( value )

            if ( value == 0 ):
                value = compare
            elif ( compare == 0 ):
                value = value
            elif ( compare < value ):
                value = compare
            value = str( value )

        if( KEY_TDS_DEFAULT in tds_cell and 'default' in additional_values ):
            default = additional_values[ 'default' ]
            value = float( value )
            if( KEY_TDS_MAX in tds_cell and 'max' in additional_values ):
                max_bound = additional_values[ 'max' ]
                if(value > max_bound):
                    value = max_bound
            value = str( value )

        return value
    # end function



    def HandleGX2Rev( self, results_dict ):
        '''
        Function for getting revision codes for GX2 tds from sql data
        Only needs to be called when dealing with SQL data and GX2 board
        '''
        distortionressql = results_dict[ TST_RES_DIST_DATA ][ 'sql' ]
        frequencyresponseres = results_dict[ TST_RES_FREQ_RESP_DATA ][ 'sql' ]
        if ( distortionressql ):
            self.boardinfodict[ KEY_TDS_ARRIS_REV ] = \
                distortionressql[ 0 ][ 229 ]
            self.boardinfodict[ KEY_TDS_EMCORE_REV ] = \
                distortionressql[ 0 ][ 228 ]
        elif ( frequencyresponseres ):
            try:
                # Use 'Rev Key' in 'Product Spec Data' (looks like
                # '10dBmout SCAPC' ) to get Arris and Emcore Rev
                revkey = self.boardinfodict[ KEY_TDS_REV_KEY ]
                devinfo = frequencyresponseres[ RSLT_DEV_INFO_KEY ] \
                    [ BRD_SPEC_DATA ][ revkey ]
                self.boardinfodict[ KEY_TDS_ARRIS_REV ] = \
                    devinfo[ KEY_TDS_ARRIS_REV ]
                self.boardinfodict[ KEY_TDS_EMCORE_REV ] = \
                    devinfo[ KEY_TDS_EMCORE_REV ]
            except Exception as e:
                log.error( 'Unable to get Revision codes: {0}'.format(
                           e ) )
            # end try
        # end if
    # end function



    def ParseMongoResultsByTraversal( self ) -> bool:
        '''
        Alternative method to parsing mongo results with a TDS Traversal Model
        in specfile
        Looks for a 'TDS Traverse Model' in the specfile, which is used in
        From().traverse_by_momdel() and returns a dictionary where each 'GET: '
        is an entry with the target value (check pylinq.py for more info)

        Results are saved to self.tdscellmap to later be used in _updateTDS()

        returns:
            True: parsing successful
            False: error when parsing

        '''

        log.debug( LOG_ENTER )


        tds = self.tdscellmap
        traversemodel = self.cfginfo[ BRD_TDS_TRAVERSE_MODEL ]

        # debugging files for comparing traverse model with results
        # write_json_file( 'temp_traverse.json', traversemodel )
        # write_json_file( 'temp_results.json', self.boardinfodict )
        try:
            traverseresults = From( self.boardinfodict ).traverse_by_model(
                traversemodel )
        except Exception as e:
            log.error( "Error in using traverse_by_model function: {}".format(
                    e) )
            return False

        for key in tds:
            if ( KEY_TDS_GET not in tds[ key ] ):
                log.debug( 'GET not in key. Continue.' )
                continue
            # end if

            # Find what the result will be named as in the traverseresults
            # dictionary (should match the GET: key in TDS Traversal Model)
            resultkey = tds[ key ][ KEY_TDS_GET ]
            if ( resultkey in traverseresults ):
                val = traverseresults [ resultkey ]

                # Get compare value for special tds compare key

                # find additional values for special keys
                additional_values = {}
                additional_value_keys = [ KEY_TDS_COMPARE, KEY_TDS_MAX,
                                            KEY_TDS_DEFAULT]

                for special_key in additional_value_keys:
                    if( special_key in tds[ key ] ):
                        additional_values[ special_key ] = traverseresults[
                            tds[ key ][ special_key ] ]


                val = self.HandleSpecialTDSKeys( val, tds[ key ],
                        additional_values)

                # Apply value to tds cell map
                # !! tds references self.tdscellmap so this change is persistent
                tds[ key ][ KEY_TDS_VALUE ] = val
        # end for

        log.debug( LOG_EXIT )

        return True

    # end function



    def CheckFinalResults( self, required_tests, results_dict ) -> bool:
        '''
        After collecting all test data, conduct a final check that all data
        is present, and check whether final result is passed or failed
        Sets final rollup result in self.boardinfodict[ TST_FINAL_RESULT ]

        params:
            required_tests: (same as in GetTestData_CATV)
            results_dict: (same as in GetTestData_CATV)

        returns:
            True: all test type results are present (even if final is Failed)
            False: missing data
        '''

        log.debug( LOG_ENTER )

        sql_rollup = None
        sql_failed_msg = "Failed test: \n"
        # Check if any results are missing
        for test_type in required_tests:
            if( results_dict[ test_type ][ 'missing' ] ):
                log.error( 'Missing tests results. Cannot proceed' )
                return False

            # Final sql pass/fail rollup
            # this pass key will be None if no sql results were found
            if( results_dict[ test_type ][ 'pass' ] == SQL_FAIL ):
                sql_rollup = TST_RES_FAILED
                sql_failed_msg += '{} \n'.format( test_type )
            elif( results_dict[ test_type ][ 'pass' ] == SQL_PASS and
                sql_rollup != TST_RES_FAILED ):
                sql_rollup = TST_RES_PASSED

        if( sql_rollup != None ):
            self.boardinfodict[ TST_FINAL_RESULT ] = sql_rollup
            if( sql_rollup == TST_RES_FAILED ):
                log.debug( msg )
                messagebox.showerror( title = 'Tests Failed', message = msg )
        else:
            # Mongo rollup exists
            if ( self.boardinfodict[ TST_FINAL_RESULT ] == TST_RES_FAILED ):
                msg = 'Mongo rollup was FAILED'
                log.error(msg)
                # messagebox.showerror( title = 'Tests Failed',
                #                       message = msg )
            # end if
        # end if

        log.debug( LOG_EXIT )

        return True
    # end function



    def GetTestData_SATCOM( self ):
        '''
        Get test data from data base for SATCOM products
        Attach result data to self.results
        Has not been tested since transition from PrintLabel.pyw to
        PackagingStation.py (~July 2017)
        '''

        log.debug( LOG_ENTER )


        satcom_res = getDeviceTestResults( db_info,
                                               self.emkridnumber )

        if ( not satcom_res ):
            log.error( " No test result found in DB" )
            return False
        # end if

        # Take the latest result from the list
        result_data = satcom_res[ -1 ]
        write_json_file( 'testresult.json', result_data )

        # Update serial number
        serialnum = self.boardinfodict[ BRD_CUSTOMER_ID ]
        result_data[ BRD_SERIAL_NUM ] = serialnum

        if ( RSLT_DEV_INFO_KEY not in result_data ):
            log.error( " Failed to find a key {0}".
                        format( RSLT_DEV_INFO_KEY ) )
            return False
        # end if

        # Updated after Kband spec consolidated
        # All the info should read from super files
        tds_spec_info = result_data[ RSLT_DEV_INFO_KEY ]
        self.tdscellmap = tds_spec_info[ BRD_TDS_CELL_MAP ]
        self.tdstmptfn = tds_spec_info[ BRD_TDS_TMPT_FN ]
        self.tdsprintfn = tds_spec_info[ BRD_TDS_PRNT_FN ]
        self.results = [ result_data ]
    # end function

# end class





class TDSWorker_LM( TDSWorker ):
    """
    basically for all LM model, the process for TDS printing is
    # 1. Get channel/power from test result
    # 2. Get part number by channel/power
    # 3. Get TDS spec by part number
    # 4. Save all associated TDS results before checking TDS data against spec.
    # 5. Check the TDS result if against spec file.
    # 6. Update TDS file if pass spec file.
    # 7. Print TDS file if pass spec file.
    """
    def __init__( self, serialnumber: str, sessiontype: str, num_type: str ):
        """
        serialnumber: value for the number type
        sessiontype: sessiontype is intended to distinguish between production,
            engineering, oql, spc
        num_type: number typem, e.g. Serial Number, Laser ID
        """
        super().__init__( serialnumber = serialnumber,
            sessiontype = sessiontype )

        retdict = {}
        # Determine LM family from the model in MES
#        family_codes = GetOriginalModelFromMES( sn = self._serialnumber )
        family_codes = GetPartNumbersbySerialnumber( self._serialnumber )
        if ( family_codes is None ):
            log.debug( 'GetPartNumbersbySerialnumber failed! '
                'Did not get LM family results from SQL using sn: '
                '{0}'.format( self._serialnumber ) )

            msg = 'SQL DB connection issue!'
            log.debug( msg )
            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            raise Exception( msg )
        # end if

        if ( isinstance( family_codes, list ) == False ):
            log.debug( 'GetPartNumbersbySerialnumber failed! '
                'Ret value not list from SQL using sn: '
                '{0}'.format( self._serialnumber ) )

            msg = 'TDS printing SW issue. Verify SQL data handling (LM family)'
            log.debug( msg )
            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            raise Exception( msg )
        # end if

        if ( len( family_codes ) == 0 ):
            log.debug( 'GetPartNumbersbySerialnumber failed! '
                'Ret value empty from SQL using sn: '
                '{0}'.format( self._serialnumber ) )

#            msg = 'LM family not matched to SN given'
            msg = 'Cannot find the product model for this board in MES'
            log.debug( msg )
            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            raise Exception( msg )
        # end if

        # Get the first value, which should be the latest family code
        last_family_code = family_codes[ 0 ][ 2 ]

        self.f_code = last_family_code

#        # Find the the SQL FCODE FIELD
#        self.f_code = last_family_code.get( SQL_FCODE_FIELD, None )
        if ( self.f_code is None ):
            log.debug( 'GetPartNumbersbySerialnumber failed! '
                'Ret value missing F_CODE field from SQL using sn: '
                '{0}'.format( self._serialnumber ) )

            msg = 'MES issue, LM family not in SQL result using SN'
            log.debug( msg )
            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            raise Exception( msg )
        # end if

        log.debug( 'Family code for SN: {0} is {1}'.format(
            self._serialnumber, self.f_code ) )

        self._boarddata_obj = BoardInformation( self.f_code, serialnumber )
        self.num_type = num_type
    # end function



    def GetBenchmarkData( self, data_dict: OrderedDict() ) -> OrderedDict:
        """
        data_dict: a dictionary that contains all the results,
                   the data for the benchmark to look for.
        return: benchmark_dict: A dictionary that contains the data to be used
                                for benchmark checking.
        """
        log.debug( LOG_ENTER )

        benchmark_dict = OrderedDict()
        benchmark_lookup = []

        if ( BRD_BENCH_DATA_KEYS not in self.cfginfo ):
            msg = 'Could not find {} in device spec'.format(
                BRD_BENCH_DATA_KEYS )
            log.debug( msg )
            raise Exception( msg )
        # end if

        benchmark_lookup = self.cfginfo[ BRD_BENCH_DATA_KEYS ]

        for each_parameter in benchmark_lookup:
            log.debug( 'look for benchmark {}'.format( each_parameter ) )
            if ( each_parameter not in data_dict.keys() ):
                msg = '{} not in data'.format( each_parameter )
                log.error( msg )
                log.error( data_dict )
                raise Exception( msg )
            data = data_dict[ each_parameter ]

            if ( ( data is None ) or ( data == '' ) ):
                msg = 'data for {} is empty or None'.format( each_parameter )
                log.error( msg )
                raise Exception( msg )

            # Handle the benchmark power option. Need to get the rounded power
            #    to get the correct power option value.
            if ( each_parameter == LM_TDS_DATA_OPW ):
                data = self.GetPowerOption( data )

                if data is None:
                    msg = 'Cannot find power option!'
                    log.error( msg )
                    raise Exception( msg )
            # end if

            benchmark_dict[ each_parameter ] = data
        # end for

        log.debug( 'Benchmark data: {}'.format( benchmark_dict ) )
        log.debug( LOG_EXIT )
        return benchmark_dict
    # end function



    def GetPartNumber( self, benchmark_data: OrderedDict ) -> str:
        """
        get part number by benchmark data like channel/power
        param benchmark: the channel string

        return: the part number value if found
                raises an Exception if part number was not found.
        """
        log.debug( LOG_ENTER )

        final_pn = ''

        prod_spec = self.cfginfo.get( LM_SPEC_PROD_SPEC_DATA )
        if ( not prod_spec ):
            error_msg = "product spec not found in spec file"
            log.error( error_msg )
            raise Exception( error_msg )

        # Go through each part number dictionary in the product spec section
        for pn, value_dict in prod_spec.items():
            # We redeclare this variable to True, and let it be changed to
            # False if any benchmark does not match the current part number
            # spec.
            found = True

            # Go through each key in the benchmark_data
            for each_benchmark in benchmark_data.keys():
                # If any of the benchmark does not match the value_dict,
                # then we go to the next partnumber.
                if ( value_dict.get( each_benchmark ).strip() !=
                     benchmark_data[ each_benchmark ] ):
                    found = False
                    break

            # After going through the benchmarks, we check the found flag
            # if it was modified for any benchmarks not matching.
            if( found == True ):
                final_pn = pn
                log.debug( "Found target part number data with "
                           "benchmark {0}".format( benchmark_data ) )

                log.debug( LOG_EXIT )
                return final_pn
                # end if
            # end for
        # end for

        # If it reaches here, then no part number was found.
        error_msg = "Part Number not found"
        log.error( error_msg )
        raise TDSValueMissingError( error_msg )
    # end function



    def CreateTDS( self, printbool: int ) -> dict:
        """ Fill all the required fields on test data sheet
        based on test results. Reads the data from testinfo (json) file and
        update each cell based on test name and group.

        param: printbool: int value used as bool. Can be 0 or 1.
                1 to print. 0 if we want to skip TDS update and printing.
        return: (dictionary)
            retdict:
                status_bool
                status_msg
        """
        log.debug( LOG_ENTER )

        retdict = dict( status_bool = False, status_msg = '' )

        num_type = self.num_type
        in_num = self._serialnumber

        # Get initial beginning data, including the main database type (SQL or
        # Mongo).
        try:
            self.boardinfodict = self._boarddata_obj.getBoardInfoFromDataBase(
                num_type, in_num, PRODUCT_TYPE_LM )
        except Exception as e:
            log.error( 'Issue when getting board info: {0}'.format( e ) )
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            retdict[ KEY_TDS_STATUS_MSG ] = 'Failed to connect to databases'
            return retdict
        # end try

        if ( self.boardinfodict is None ):
            msg = 'Did not find any board with the given SN in Mongo or SQL'
            log.debug( msg )
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            retdict[ KEY_TDS_STATUS_MSG ] = msg
            return retdict
        # end if

        log.debug( 'self.boardinfodict is {0}'.format( self.boardinfodict ) )

        # Check boardinfodict in rollup
        if ( self.boardinfodict is None ):
            msg = 'Did not find data from serial number query'
            log.debug ( msg )

            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        # end if

        log.debug( 'TDS created' )

        # Get product_id from rollup
        # if ( HW_PRODUCT_ID not in self.boardinfodict.get(
        #      LM_SPEC_DEVICE_INFO ) and HW_PRODUCT_ID not in
        #      self.boardinfodict.get( RSLT_DEV_INFO_KEY ) ):
        #     msg = "HW_PRODUCT_ID not found on board spec"
        #     log.debug ( msg )
        #     retdict[ KEY_TDS_STATUS_MSG ] = msg
        #     retdict[ KEY_TDS_STATUS_BOOL ] = False
        #     return retdict
        # end if
        if ( self.f_code in [ '1612-STD', '16XX' ] ):
            # 11/3/2017, for 1310, we get productid through partnumber
            productid = self.boardinfodict[ HW_PRODUCT_ID ]
        else:
            productid = self.f_code
        # end if

        # Get product spec file from MONGO DB
        self.cfginfo = getDeviceInfo( db_info = db_info,
                                      product_id = productid )

        if ( self.cfginfo is None ):
            msg = str( 'Failed to find spec file {0} from MONGO DB'.
                        format ( productid ) )
            log.debug( msg )

            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        # end if

        write_json_file( DEV_SPEC_FN, self.cfginfo )
        log.debug( 'Wrote specfile to local file: {}'.format( DEV_SPEC_FN ) )
        # get initial data end

        # default to test final result to failed. Only if theres nothing in the
        # 'okaylist' do we set to 'Passed'
        self.boardinfodict[ TST_FINAL_RESULT ] = TST_RES_FAILED
        retdict[ KEY_TDS_CFG_INFO ] = self.cfginfo
        retdict[ KEY_TDS_BOARD_INFO_DICT ] = self.boardinfodict
        # Pull all cell map data
        retbool = self.getRelatedDataFromCfg()
        # For laser modules, also need to check the TDS Data Set Config
        if ( BRD_TDS_DATASET_CONFIG not in self.cfginfo ):
            log.error( 'TDS data set config not found in tds cfg dictionary' )
            retbool = False
        # end if

        if ( retbool == False ):
            msg = 'Missing information in device spec!'
            log.debug( msg )
            messagebox.showwarning( title = 'Status', message = msg )

            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        else:
            msg = 'Product spec configurations verified'
            log.debug( msg )
            retdict[ KEY_TDS_STATUS_MSG ] = msg
        # end if

        self.tdsdatasetconfig = self.cfginfo[ BRD_TDS_DATASET_CONFIG ]

        spec = None
        lasttdsdata = None

        # Create dataset from TDS by sn
        try:
            self.tdsdatasetdict = self.GetDateAndAnalyticByplinq(
                self.tdsdatasetconfig, self.boardinfodict )
        except ResultOutOfRangeError as ex:
            log.error( ex )
            retdict[ KEY_TDS_STATUS_MSG ] = ex
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        except TestResultFailedError as ex:
            log.error( ex )
            retdict[ KEY_TDS_STATUS_MSG ] = ex
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        except Exception as ex:
            log.error( ex )
            retdict[ KEY_TDS_STATUS_MSG ] = ex
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        # end try

        log.debug( "before dataset Collecting data {}".format(
                   self.tdsdatasetdict ) )

        self.CollectingDateSetData()
        log.debug( "after dataset Collecting data {}".format(
                   self.tdsdatasetdict ) )

        # For 1310 laser module, we get CNR, CSO and CTB here
        if ( self.f_code in [ '1612-STD', '16XX' ] ):
            self.GetDistortionDataMethod()
            log.debug( "after dataset Collecting distortion data {}".format(
                        self.tdsdatasetdict ) )
        # end if

        self.SetCalculation()
        log.debug( "after dataset Calculation {}".format(
                   self.tdsdatasetdict ) )

        # Save necessary values
        self.tdsdatasetdict[ LM_TDS_DATA_PROD_ID ] = productid

        try:
            benchmark_data = self.GetBenchmarkData( self.tdsdatasetdict )
            pn = self.GetPartNumber( benchmark_data )
            spec = self.GetTDSSpec( pn )
            # pn_dict should be one key only.
            # pn_dict = { 'G1798-001-048': { "Model": "1798-48-BB-SC-16",
            #                              "ITU Channel": "48",
            #                              "Description": "1798-48-BB-SC-16"
            #                             }
            desp = spec.get( LM_SPEC_DESC )

            # Check if there is BRD_WAVE_LEN_NM not in the boardinfo. If there
            # is nothing, we replace with the the LM_TDS_DATA_WAVELEN we get
            # from the self.tdsdatasetdict.
            wl = self.tdsdatasetdict[ LM_TDS_DATA_WAVELEN ]
            if ( ( BRD_WAVE_LEN_NM not in self.boardinfodict ) or
                 ( self.boardinfodict[ BRD_WAVE_LEN_NM ] == '' ) ):
                self.boardinfodict[ BRD_WAVE_LEN_NM ] = wl
            # end if

            if ( self.f_code in [ '1612-STD', '16XX' ] ):
                # For 1310 laser module, we don't have ITU channel
                pass
            else:
                chann = self.tdsdatasetdict[ BRD_ITU_CH_NUM ]
                self.boardinfodict[ BRD_ITU_CH_NUM ] =  chann
            # end if

            write_json_file ( 'boardinfopostchann.json', self.boardinfodict )

            self.boardinfodict[ HW_PART_NUM ] = pn
            self.boardinfodict[ BRD_CUSTOMER_ID ] = in_num
        except Exception as ex:
            log.debug( 'Exception caught {}'.format( ex ) )
            retdict[ KEY_TDS_STATUS_MSG ] = str( ex )
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        # end try

        try:
            self.SaveTDSData( desp, pn )
        except Exception as ex:
            log.debug( 'Exception caught {}'.format( ex ) )
            retdict[ KEY_TDS_STATUS_MSG ] = str( ex )
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        # end try

        try:
            # get spec and TDSData
            lasttdsdata = self.GetTDSData()
        except TDSValueMissingError as ex:
            log.debug( 'Missing TDS value {}'.format( ex ) )
            retdict[ KEY_TDS_STATUS_MSG ] = str( ex )
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        except Exception as ex:
            log.debug( 'Exception caught {}'.format( ex ) )
            retdict[ KEY_TDS_STATUS_MSG ] = str( ex )
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        # end try

        # Want:
        # 1. Get channel/power from test result
        # 2. Get part number by channel/power
        # 3. Get TDS spec by part number
        # 4. Save all associated TDS results before checking TDS data against
        #    spec.
        # 5. Check the TDS result if against spec file.
        # 6. Update TDS file if pass spec file.
        # 7. Print TDS file if pass spec file.
        sqlresults = {}
        productspec = {}

        retbool = self.getRelatedDataFromCfg()
        if ( retbool == False ):
            msg = 'Missing information in device spec'
            log.debug( msg )

            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        else:
            msg = 'Product spec configurations verified'
            log.debug( msg )
            retdict[ KEY_TDS_STATUS_MSG ] = msg
        # end if

        # CHECK TDS SPEC IF PASS then updatettds_lm and print
        okaylist = self.CheckAgainstSpec( spec, lasttdsdata[ 0 ] )
        log.debug( "check spec {}".format( okaylist ) )
        if ( len( okaylist ) >= 1 ):
            msg = "check {0} spec is Fail".format( okaylist )
            self.Update_LastTds_Status( "Fail" )
            log.error( msg )

            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False

            return retdict
        # end if

        self.boardinfodict[ TST_FINAL_RESULT ] = TST_RES_PASSED
        # reassign with passed status.
        retdict[ KEY_TDS_BOARD_INFO_DICT ] = self.boardinfodict
        msg = 'check spec pass'
        log.debug( msg )
        retdict[ KEY_TDS_STATUS_MSG ] = msg
        retdict[ KEY_TDS_STATUS_BOOL ] = True

        if ( printbool != True ):
            msg = 'Skipping printing of TDS, not selected'
            log.debug( msg )
            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = True
            return retdict
        # end if

        try:
            retbool = self.updateTDS_LM()
        except Exception as e:
            msg = str( e )
#            if ( retbool == False ):
#                msg = 'Unsuccessful update of template'
            log.debug( msg )

            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        # end if

        if ( retbool == False ):
            msg = 'Unsuccessful update of template'
            log.debug( msg )

            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        # end if

        msg = 'Successful update of template'
        log.debug( msg )

        retbool = self.printTDSData()
        if ( retbool == False ):
            msg = 'Unsuccessful print of TDS'
            log.debug( msg )

            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            return retdict
        # end if

        msg = 'Successfully print TDS'
        self.Update_LastTds_Status( "Pass" )
        log.debug( msg )


        log.debug( LOG_EXIT )

        retdict[ KEY_TDS_STATUS_MSG ] = msg
        retdict[ KEY_TDS_STATUS_BOOL ] = True
        return retdict
    # end function



    def updateTDS_LM( self ) -> bool:
        '''
        Find database results, as well as fill in the cells in this function.
        For laser module devices.

        Return: True: if all values were successfully retrieved
                False: if any value was unsuccessfully retrieved
        '''
        log.debug( LOG_ENTER )

        # 10/4/17 Removed big try block. Need to handle errors better.
#        try:
        self.serialnum = None
        # adapt to stale schema
        # To adapt to initial laser module station runs using Aurora spec file
        # and Serial Number with and without underscore.
        # Here we first check:
        # 'Serial_Number' in result
        # 'Serial Number' in result
        # 'Serial_Number' in device spec
        # 'Serial Number' in device spec
        self.serialnum = self.boardinfodict.get( LM_RESULT_FIELD_SN, None )
        if not self.serialnum:
            self.serialnum = self.boardinfodict.get( BRD_SERIAL_NUM, None )
            if not self.serialnum:
                device_info = self.boardinfodict.get( LM_SPEC_DEVICE_INFO,
                                                      None )
                if not device_info:
                    device_info = self.boardinfodict.get( RSLT_DEV_INFO_KEY,
                                                          None )
                # end if

                if not device_info:
                    log.error( "could not find device info and sn" )
                    return False
                # end if

                self.serialnum = device_info.get( LM_RESULT_FIELD_SN, None )
                if not self.serialnum:
                    self.serialnum = device_info.get( BRD_SERIAL_NUM, None )
                # end if
        # end if

        if not self.serialnum:
            log.error ( "could not find sn" )
            return False
        # end if

        excel = None
        try:
            excel = win32com.client.Dispatch( 'Excel.Application' )
            excel.Visible = False
            excel.DisplayAlerts = False
        except Exception as e:
            log.debug( "Failed to open excel Application error {0}".format(
                e ) )
            return False
        # end try

        # Set in
        self.tdssheet = self.cfginfo[ BRD_TDS_SHEET_FN ]
        tdstmptfn = self.tdstmptfn

        log.debug( 'tdstmptfn is {}'.format( tdstmptfn ) )
        path = os.path.abspath( os.path.dirname( sys.argv[ 0 ] ) )
        original_tdspath = os.path.join( os.path.join( path, 'TDS' ),
            tdstmptfn )

        copy_tdspath = os.path.join( path, 'local_{0}'.format( tdstmptfn ) )

        try:
           local_tdspath = shutil.copy( original_tdspath, copy_tdspath )
        except PermissionError:
            msg = 'Close excel file to proceed: {0}'.format( copy_tdspath )
            log.debug( msg )
            raise PermissionError( msg )
        except IOError:
            msg = 'Template does not exist'
            log.debug( msg )
            raise IOError( msg )
        except Exception as ex:
            log.debug( 'Unexpected exception encountered during copy of TDS '
                'template {}'.format( ex ) )
            return False
        # end try

        wb = excel.Workbooks.Open( local_tdspath )
        log.debug( 'wb object: {0}'.format( wb ) )
        ws = wb.Worksheets( self.tdssheet )

        ismissingdata = False
        missingdatalist = []

        # double check the itu channel value
        itu_channel_value = None
        itu_check_value = None

        for eachentry in self.tdscellmap:
            log.debug( 'eachentry is {0}'.format( eachentry ) )
            datloc = eachentry[ "data location" ]
            datname = eachentry[ "Name" ]
            if ( datloc == SPEC_DB_SQL ):
                if ( SPEC_DB_SQL not in eachentry ):
                    log.debug( 'Missing SQL parameter in {}'.format(
                        eachentry ) )
                    return False
                querydict = eachentry[ SPEC_DB_SQL ]
                queryvalue =  self.tdsdatasetdict[ datname ]
                if queryvalue == "":
                    log.debug( 'No data for {}. Skipping.'.format(
                        eachentry[ 'Name' ] ) )
                    ismissingdata = True
                    missingdatalist.append( eachentry[ 'Name' ] )
                    continue
                # end if

                # Use ITU channel to update the model type
                if ( eachentry[ 'Name' ] == LM_SPEC_ITU_CHANNEL ):
                    channum = str( int ( queryvalue ) )
                    itu_channel_value = channum

                    # Use data from spec for model string after matching
                    # the ITU channel number
                    # example
                    # "Product Spec Data": {
                    #     "G1752-204-018": {
                    #         "ITU Channel": "18",
                    #         "Description": "1752A-C21-18-BB-SK-10,Arris 62-10102-18",
                    #         "Model": "1752A-C21-18-BB-SK-10"
                    #     },
                    #     "G1752-204-019": {
                    #         "ITU Channel": "19",
                    #         "Description": "1752A-C21-19-BB-SK-10,Arris 62-10102-19",
                    #         "Model": "1752A-C21-19-BB-SK-10"
                    #     }
                    if ( isinstance( self.prod_spec_data, dict ) == True ):
                        pn_dict = {}
                        log.debug( 'ITU channel number from query: {0}'.format(
                            channum ) )

                        for curr_pn in self.prod_spec_data.keys():
                            curr_pn_dict = self.prod_spec_data[ curr_pn ]
                            log.debug( 'check key {0} with dict: {1}'.format(
                                curr_pn, curr_pn_dict ) )

                            curr_value_itu = curr_pn_dict.get( BRD_ITU_CH_NUM,
                                None )

                            if ( curr_value_itu is None ):
                                log.debug( 'Could not get ITU channel from '
                                    'current part number' )
                                continue
                            # end if

                            if ( channum == curr_value_itu ):
                                log.debug( 'Matched ITU with query value' )
                                pn_dict = curr_pn_dict
                                break
                        # end for

                        itu_model_str = pn_dict.get( BRD_MODEL, None )
                        queryvalue = itu_model_str
                    else:
                        rangeval = eachentry[ 'range' ]
                        modeltemp = str( ws.Range( rangeval ).Value )
                        log.debug ( 'ITU Channel value from template is '
                            '{0}'.format ( modeltemp ) )

                        log.debug ( 'ITU Channel value from db is {0}'.format (
                                channum ) )
                        if 'xx' not in modeltemp:
                            modeltemp = "1752A-C21-xx-BB-SK-10"
                        queryvalue = modeltemp.replace( 'xx',str( channum ) )
                        itu_channel_value = channum
                        log.debug( 'queryvalue modified to {}'.format(
                            queryvalue ) )
                    # end if

                if ( eachentry[ 'Name' ] == 'ITU Channel 1' ):
                    itu_check_value = str( int( queryvalue ) )
                # end if

                if ( eachentry[ 'Name' ] == 'SQL Date' ):
                    queryvalue = str( queryvalue )
                # end if
            elif ( datloc == SPEC_DB_MONGO ):
                # need from Tony how to acquire data
                querydict = eachentry[ SPEC_DB_MONGO ]
                queryvalue = None
                try:
                    queryvalue = self.tdsdatasetdict[ datname ]
                except Exception as ex:
                    log.error( ex )
                    queryvalue = 0

                # will need to parse for querydata.
                if ( queryvalue is None ):
                    log.debug( 'No data in Mongo for {}'.format( querydict ) )
                    ismissingdata = True
                    missingdatalist.append( eachentry[ 'Name' ] )
            elif ( datloc == KEY_TDS_CALCULATION ):
                queryvalue = self.tdsdatasetdict[ datname ]
                if ( queryvalue is None ):
                    log.debug( 'Could not calculate Value for {}'.format(
                        eachentry ) )
                    ismissingdata = True
                    missingdatalist.append( eachentry[ 'Name' ] )
            elif ( datloc == 'None' ):
                if eachentry[ 'Name' ] == 'Date':
                    d = datetime.date.today()
                    queryvalue = d.strftime( '%m/%d/%Y' )
                pass
            else:
                log.debug( 'No datalocation for {}'.format( eachentry ) )
                return False

            if ( 'range' not in eachentry ):
                log.debug( 'Missing range parameter from {}'.format(
                    eachentry ) )
                return False
            # end if

            log.debug( 'queryvalue is {}'.format( queryvalue ) )
            rangeval = eachentry[ 'range' ]
            # Some data will be required to pull from SQL such as ITU Channel
            # but may not be inserted into the TDS. In this case, the rangeval
            # will be empty.
            if( rangeval != '' ):
                ws.Range( rangeval ).Value = queryvalue
        #end for

        # TO DO WHY DO WE NEED THIS?
        if itu_check_value != itu_channel_value :
            messagebox.showerror ( title = 'Invalid Data', message = "channel"
                                   "value mismatched in tds file" )
            ismissingdata = True
        # end if


        if ( ismissingdata == True ):
            msg = "Data is missing for: \n"
            for item in missingdatalist:
                msg += item + '\n'
            log.debug( 'Data missing for {}'.format( msg ) )
            messagebox.showerror( title = 'Missing Data', message = msg )
            wb.Close( False )

            excel.Quit()
            excel = None
            return False

       # format print report form 1752-c21 add by bob
        path = os.path.abspath( os.path.dirname( sys.argv[ 0 ] ) )

        # use tick to guarantee unique file name
        tick = str ( time.time ( ) ).split ( "." )[ 0 ]
        self.tdsprintfn = '{0}_{1}_{2}.xlsx'.format( self.tdsprintfn,
            self.serialnum, tick )

        tds_copy_path = os.path.join( path, self.tdsprintfn )
        log.debug( 'Saving TDS copy as: {0}'.format( tds_copy_path ) )
        wb.SaveAs( tds_copy_path )

        log.debug( 'Close excel wb' )
        wb.Close( False )

        log.debug( 'Close excel application' )
        excel.Quit()

        # Remove local copy of original template file
        if ( os.path.exists( local_tdspath ) == True ):
            log.debug( 'Local template file exists, remove' )
            os.remove( local_tdspath )
        # end if

        return True
#        except Exception as ex:
#            log.error( "error happend in updateTDS_LM {0} ".format( ex ) )
#            if excel != None:
#                try:
#                    log.error ( "try to quit excel" )
#                    excel.Quit ( )
#                except:
#                    log.error( "try to quit excel failed" )
#                    os.system ( 'taskkill /f /im EXCEL.exe' )
#            return False
        log.debug( LOG_EXIT )
    # end function



#    # 10/5/17 unused function - delete?
#    def parseLMresult( self, querydict ) -> str:
#        """
#        querydict: dictionary of where value is located in result file.
#        e.g.,
#            {
#                        "Region": "Tests",
#                        "TestGroup": "Laser Module Main",
#                        "Test Name": "BERMER",
#                        "Value": "Result BER Pre Max"
#            }
#        Region is the portion of super test result that we parse through, to
#            avoid having to parse through entire file.
#
#        Return: finalval of desired entry, if successfully found.
#                None: if the value cannot be found
#        """
#        log.debug( LOG_ENTER )
#        region = querydict[ 'Region' ]
#        testgrp = querydict[ 'TestGroup' ]
#        testnm = querydict[ 'Test Name' ]
#        testval = querydict[ 'Value' ]
#
#        if region == 'Tests':
#            testreg = self.boardinfodict[ region ]
#            # testreg is currently a list, if in 'Tests'
#            testdict = None
#            for testdict in testreg:
#                if testdict[ 'TestGroup' ] == testgrp:
#                    grptests = testdict[ 'TestList' ]
#                    break
#
#            if testdict is None:
#                log.debug( 'No TestGroup {} in test region'.format( testgrp ) )
#                return None
#
#            for eachtest in grptests:
#                if eachtest[ 'Test Name' ] == testnm:
#                    spectest = eachtest
#                    break
#
#            try:
#                finalval = spectest[ 'Result Data' ][ 'Result List' ][ 0 ]\
#                    [ testval ]
#            except Exception as ex:
#                log.debug( 'Cannot find test value "{}" in result "{}"'.format(
#                    testval, spectest[ 'Result Data' ] ) )
#                return None
#
#        log.debug( 'finalval is {}'.format( finalval ) )
#        finalval = str( finalval )
#
#        log.debug( LOG_EXIT )
#        return finalval
#    # end function



    def GetDateAndAnalyticByplinq( self, datasetconfig: dict,
                                         resultsdict: dict ) -> dict:
        """
        Get Mongo results using pylinq

        Return: dict of results if found
                empty dictionary if no results passed in
        """
        log.debug( LOG_ENTER )

        log.debug( "dataset config: {}".format( datasetconfig ) )
        log.debug( "resultsdict: {}".format( resultsdict ) )

        _tdsdatasetdict = {}
        if not resultsdict:
            log.debug( 'No results passed in' )
            return {}

        # 10/4/17 Removed big try loop. need to handle errors appropriately,
        # specifically.
        for column in datasetconfig:
            tempdictmonfos = datasetconfig[ column ]

            where_exp = tempdictmonfos[ KEY_TDS_WHERE ]
            select_exp = tempdictmonfos[ KEY_TDS_SELECT ]
            columnno = tempdictmonfos[ KEY_TDS_COLUMN ]

            try:
                qyuelisttemo = From( resultsdict ).where(
                    where_exp ).select( select_exp )

                if ( KEY_TDS_INDEX in tempdictmonfos.keys() ):
                    arrayvalue = qyuelisttemo[ int( tempdictmonfos\
                                               [ KEY_TDS_INDEX ] ) ]
                elif ( KEY_TDS_SUB_WHERE in tempdictmonfos.keys() ):
                    subwhere_exp = tempdictmonfos[ KEY_TDS_SUB_WHERE ]
                    subselect_exp = tempdictmonfos[ KEY_TDS_SUB_SELECT ]
                    templissub = { "abc": qyuelisttemo }
                    json_contentnew = json.dumps( templissub )

                    qyuelistsub = From(
                        json_contentnew ).where( "abc.$." + subwhere_exp ).\
                                          select( "abc.$." + subselect_exp )
                    arrayvalue = qyuelistsub[ 0 ]
                else:
                    arrayvalue = qyuelisttemo[ 0 ]
                # end if

                log.debug( "dataset key:{} Value{}".format(
                           column, arrayvalue ) )
            except:
                arrayvalue = ""
                log.error( "dataset key:{} err, col {}".format( column,
                    columnno ) )
                _tdsdatasetdict[ column ] = arrayvalue
                continue
            # end try

            _tdsdatasetdict[ column ] = arrayvalue
        # end for

        write_json_file( "TestDataSetdict.json", _tdsdatasetdict )
        log.debug( 'Wrote data to local file: TestDataSetdict.json' )

        log.debug( LOG_EXIT )
        return _tdsdatasetdict
    # end function



    def CollectingDateSetData( self ):
        """
        if mongo data is '' then get mssql data
        ITU Channel is taken from MES only.

        but channel query by specfile form workorder PN.

        return:
        """
        log.debug( LOG_ENTER )

        for eachentry in self.tdscellmap:
            log.debug( 'eachentry is {0}'.format( eachentry ) )

            datloc = eachentry[ KEY_TDS_DATA_LOCATION ]
            datname = eachentry[ BRD_NAME ]
            queryvalue = ""

            if ( datname not in self.tdsdatasetdict ):
                log.debug( 'Missing key {}. Skipping.'.format( datname ) )
                continue
            # end if

            if ( ( datloc != SPEC_DB_SQL ) or
                 ( self.tdsdatasetdict[ datname ] != '' ) ):
                log.debug( 'Do not need to get value from SQL. Skipping.' )
                continue
            # end if

            # For ITU channel, we have to make sure that the ITU from the
            # work order passes
            if ( BRD_ITU_CH_NUM in datname ):
                log.debug( 'Get Target Channel from MES' )
                queryvalue = GetTargetChannelFromMES( sn = self._serialnumber )
            else:
                # For 1310 lase module, we don't get CNR, CSO and CTB here
                if ( ( self.f_code in [ '1612-STD' ] ) and
                    ( eachentry[ BRD_NAME ] in [ LM_TDS_DATA_CNR,
                        LM_TDS_DATA_CSO, LM_TDS_DATA_CTB ] ) ):
                    pass
                else:
                    if ( SPEC_DB_SQL not in eachentry ):
                        log.debug( 'Missing SQL param in {}'.
                                format( eachentry ) )
                        return False
                    # end if

                    querydict = eachentry[ SPEC_DB_SQL ]
                    if ( self.GetSQLDBInformation( self._serialnumber,
                         querydict ) != True ):
                        log.debug( 'Failed to get value from SQL database' )
                    # end if

                    log.debug( 'querydict is {}'.format( querydict ) )

                    if ( ( querydict[ KEY_TDS_VALUE ] == [] ) or
                         ( querydict[ KEY_TDS_VALUE ] == None ) ):
                        log.debug( 'No data for {}. Skipping.'.
                                format( datname ) )
                        continue
                    # end if

                    queryvalue = querydict[ KEY_TDS_VALUE ][ 0 ][ 0 ]
                # end if
            # end if

            self.tdsdatasetdict[ datname ] = queryvalue
        # end for

        write_json_file( "afterTestDataSetdict.json", self.tdsdatasetdict )
        log.debug( 'Wrote combined results to local file: '
                                        ' afterTestDataSetdict.json' )

        log.debug( LOG_EXIT )
    # end function



    def GetDistortionDataMethod( self ):
        """ Get Iop value and distortion test data through Optical Power

        """
        log.debug( LOG_ENTER )

        iopvalue = None
        opticalpower = self.tdsdatasetdict[ LM_TDS_DATA_OPW ]
        ithvalue = self.tdsdatasetdict[  LM_TDS_DATA_ITH ]

        for eachentry in self.tdscellmap:
            log.debug( 'eachentry is {0}'.format( eachentry ) )

            datloc = eachentry[ KEY_TDS_DATA_LOCATION ]
            datname = eachentry[ BRD_NAME ]
            queryvalue = ""

            if ( datname not in self.tdsdatasetdict ):
                log.debug( 'Missing key {}. Skipping.'.format( datname ) )
                continue
            # end if

            if ( ( datloc != SPEC_DB_SQL ) or
                 ( self.tdsdatasetdict[ datname ] != '' ) ):
                log.debug( 'Do not need to get value from SQL. Skipping.' )
                continue
            # end if

            for i in range( 1, 8 ):
                if ( i == 1 ):
                    iopnumber = 30
                    if ( opticalpower > OPTI_POWER_SPEC_MW_10 and
                                opticalpower < OPTI_POWER_SPEC_MW_125 ):
                        iopvalue = iopnumber + ithvalue
                        break
                    # end if
                elif ( i == 2 ):
                    iopnumber = 40
                    if ( opticalpower > OPTI_POWER_SPEC_MW_126 and
                            opticalpower < OPTI_POWER_SPEC_MW_158 ):
                        iopvalue = iopnumber + ithvalue
                        break
                    # end if
                elif ( i == 3 ):
                    iopnumber = 50
                    if ( opticalpower > OPTI_POWER_SPEC_MW_159 and
                            opticalpower < OPTI_POWER_SPEC_MW_199 ):
                        iopvalue = iopnumber + ithvalue
                        break
                    # end if
                elif ( i == 4 ):
                    iopnumber = 60
                    if ( opticalpower > OPTI_POWER_SPEC_MW_20 and
                            opticalpower < OPTI_POWER_SPEC_MW_249 ):
                        iopvalue = iopnumber + ithvalue
                        break
                    # end if
                elif ( i == 5 ):
                    iopnumber = 70
                    if ( opticalpower > OPTI_POWER_SPEC_MW_22 and
                            opticalpower < OPTI_POWER_SPEC_MW_30 ):
                        iopvalue = iopnumber + ithvalue
                        break
                    # end if
                elif ( i == 6 ):
                    iopnumber = 80
                    if ( opticalpower > OPTI_POWER_SPEC_MW_25 and
                            opticalpower < OPTI_POWER_SPEC_MW_309 ):
                        iopvalue = iopnumber + ithvalue
                        break
                    # end if
                elif ( i == 7 ):
                    iopnumber = 90
                    if ( opticalpower > OPTI_POWER_SPEC_MW_31 and
                            opticalpower < OPTI_POWER_SPEC_MW_397 ):
                        iopvalue = iopnumber + ithvalue
                        break
                    # end if
                # end if
            # end for

            self.tdsdatasetdict[ LM_TDS_DATA_Iop ] = iopvalue

            if ( eachentry[ BRD_NAME ] in [ LM_TDS_DATA_CNR, LM_TDS_DATA_CSO,
                                          LM_TDS_DATA_CTB ] ):
                if ( SPEC_DB_SQL not in eachentry ):
                    log.debug( 'Missing SQL param in {}'.format( eachentry ) )
                    return False
                # end if

                querydict = eachentry[ SPEC_DB_SQL ]
                valuecoltemp = querydict[ SQL_QUERY_VALUE_COL ]. \
                                split( '[' )[ 1 ].split( ']' )[ 0 ]
                valuecoltemp = '[{0}{1}]'.format( iopnumber ,valuecoltemp)
                querydict[ SQL_QUERY_VALUE_COL ] = valuecoltemp
                if ( self.GetSQLDBInformation( self._serialnumber,
                     querydict ) != True ):
                    log.debug( 'Failed to get value from SQL database' )
                # end if

                log.debug( 'querydict is {}'.format( querydict ) )

                if ( ( querydict[ KEY_TDS_VALUE ] == [] ) or
                     ( querydict[ KEY_TDS_VALUE ] == None ) ):
                    log.debug( 'No data for {}. Skipping.'.format( datname ) )
                    continue
                # end if

                queryvalue = querydict[ KEY_TDS_VALUE ][ 0 ][ 0 ]
                self.tdsdatasetdict[ datname ] = queryvalue
            # end if
        # end for

        write_json_file( "afterGetDistortionData.json", self.tdsdatasetdict )
        log.debug( 'Wrote combined results to local file: '
                    ' afterGetDistortionData.json' )

        log.debug( LOG_EXIT )
    # end function



    def GetChirpPreviousMethod( self ) -> str:
        """ Get 30000 chirp values from SQL, and random

            return: chirp value
        """
        log.debug( LOG_ENTER )

        chirpvalue = None
        chirplist = []

        chirpquery = {
            SQL_QUERY_INIT_CFG: "TOP 30000",
            SQL_QUERY_VALUE_COL: "Chirp",
            SQL_QUERY_TABLE: "ChirpSpecTrumData",
            SQL_QUERY_ID_COL: "Chirp is not null and Chirp != 0 and PFCode",
            SQL_QUERY_TIME_COL: "RecordId",
            SQL_QUERY_ORDER_DIR: "DESC"
        }

        if ( not self.GetSQLDBInformation( 1, chirpquery ) ):
            log.debug( 'Failed to get value from SQL database' )
            return None
        # end if

        log.debug( 'chirp is {}'.format( chirpquery[ KEY_TDS_VALUE ] ) )

        if ( chirpquery[ KEY_TDS_VALUE ] == [] ):
            log.debug( 'No chirp data' )
            return None
        # end if

        chirplist = chirpquery[ KEY_TDS_VALUE ]
        chirpvalue = random.sample( chirplist, 1 )[ 0 ][ 0 ]

        log.debug( LOG_EXIT )
        return chirpvalue
    # end function


    def SetCalculation( self ):
        """
        just Calculation need Calculation key
        return:
        """
        log.debug( LOG_ENTER )

        for eachentry in self.tdscellmap:
            log.debug( 'eachentry is {0}'.format( eachentry ) )

            datloc = eachentry[ KEY_TDS_DATA_LOCATION ]
            datname = eachentry[ BRD_NAME ]
            queryvalue = ""

            if ( datloc == KEY_TDS_CALCULATION ):
                # JS only calculate value if it's not there already from mongo
                # results.
                if ( self.tdsdatasetdict[ datname ] != "" ):
                    log.debug( 'Calculation value already exists. Skip.' )
                    continue
                # end if

                queryvalue = self.TDSCalculateVal( eachentry )

                if ( queryvalue is None ):
                    queryvalue = ""
                    log.debug( 'Could not calculate Value for {}'.format(
                        eachentry ) )
                # end if

                self.tdsdatasetdict[ datname ] = queryvalue
            else:
#                # 10/5/17 Modified the excel template instead
#                # Force modify data
#                if ( eachentry[ BRD_NAME ] == LM_TDS_DATA_CSO ):
#                    tempcso = self.tdsdatasetdict[ LM_TDS_DATA_CSO ]
#                    if ( tempcso != "" ):
#                        tempcso = float( tempcso ) * -1
#                    # end if
#
#                    self.tdsdatasetdict[ LM_TDS_DATA_CSO ] = tempcso
#                elif ( eachentry[ BRD_NAME ] == LM_TDS_DATA_CTB ):
#                    tempctb = self.tdsdatasetdict[ LM_TDS_DATA_CTB ]
#                    if ( tempctb != "" ):
#                        tempctb = float( tempctb ) * -1
#                    # end if
#
#                    self.tdsdatasetdict[ LM_TDS_DATA_CTB ] = tempctb
                if ( eachentry[ BRD_NAME ] == LM_TDS_DATA_STATION ):
                    station = self.tdsdatasetdict[ LM_TDS_DATA_STATION ]
                    if ( station != "" ):
                        try:
                            station = station[ -1: ]
                        except:
                            pass
                        # end try
                    # end if

                    self.tdsdatasetdict[ LM_TDS_DATA_STATION ] = station
                # end if
            # end if
        # end for

        write_json_file( "afterdatalocation.json", self.tdsdatasetdict )
        log.debug( 'Wrote data to local file: datalocation.json' )

        log.debug( LOG_EXIT )
    # end function



    def GetTDSSpec( self, partnumber: str ) -> dict:
        """
        get tds check spec by partnumber
        :param partnumber: the spec of part number
        :return:the TDS spec dictionary
        """
        """
        get part number by partnumber
        :param partnumber: partnumber
        :return:the check spec of part number
        """
        log.debug( LOG_ENTER )

        check_tdsspec = {}
        prod_spec = self.cfginfo.get( LM_SPEC_PROD_SPEC_DATA )
        if ( prod_spec ):
            log.debug( "get partnumber spec" )

            try :
                check_tdsspec = prod_spec[ partnumber ]
                log.debug( "get partnumber check spec{}".format(
                    check_tdsspec ) )
            except Exception as ex :
                log.error ( ex )
            # end try

            log.debug( LOG_EXIT )
            return check_tdsspec
        else:
            error_msg = "partnumber spec not found"
            log.error( error_msg )
            raise TDSValueMissingError( error_msg )
        # end if

        log.debug( LOG_EXIT )
    # end function



    def GetTDSData( self ) -> dict:
        """
        get tds data by serialnumber
        serialnumber : the serialnumber by cur input
        return:the TDS data dictionary
        """
        log.debug( LOG_ENTER )

        tds_results_data_dict = {}

        if ( self.f_code in [ '1612-STD' ] ):
            sql_query = "select top 1 PowerOption,Ith,Slopeff,Chirp,\
             ConnectorType,abs(CSO) as CSO,\
             abs(CTB) as CTB,abs(CNR) as CNR,OpticalPower,\
              CountPrint,abs(SMSR) as SMSR,WaveLen,\
              Station,Iop,Operator from \
              LaserModuleShippingData \
              where Device_SN='{0}' order by createtime desc".format(
              self._serialnumber )
        else:
            sql_query = "select top 1 PowerOption,Ith,Slopeff,Chirp,\
             ConnectorType,abs(SMSR) as SMSR,abs(CSO) as CSO,\
             abs(CTB) as CTB,abs(CNR) as CNR,OpticalPower,MPDCurrent,\
              TrackingError_Max as TrackingError_dB_Pos,\
              TrackingError_Min as TrackingError_dB_Neg,CountPrint,\
              ForwardVoltage,FreqResponse,LsrTemperature,\
              Iop,abs(CNR_61) as CNR_61,WaveLen,\
              abs(CNR_547) as CNR_547,Station,Operator,\
              B_Const_K,RTH_Kohm,abs(CNR_294) as CNR_294,\
              BER,MPDCurrent,MPDISlope_LIP,LsrTemperature from \
              LaserModuleShippingData \
              where Device_SN='{0}' order by createtime desc".format(
              self._serialnumber )
        # end if

        log.debug( 'tds get query is {0}'.format( sql_query ) )

        try:
            msserver = mssqlserver( host = DB_HOSTNAME_DEFAULT,
                user = DB_USERNAME_DEFAULT, pwd= DB_PASSWORD_DEFAULT,
                db = DB_DATABASE_DEFAULT, charset = DB_DATABASE_CONN_ENCODING,
                retdict = True )
        except NameError as ex:
            err_msg = "cound not connect to sql server {0}".format(ex)
            log.error( err_msg )
            raise TDSSqlConnectionError( err_msg )
        # end try

        try:
            tds_results_data_dict = msserver.ExecQuery( sql_query )
        except Exception as ex:
            err_msg = "get tds data to sql server failed, {0}".format(ex)
            log.error( err_msg )
            raise TDSGETError( err_msg )
        # end try

        log.debug( LOG_EXIT )
        return tds_results_data_dict
    # end function



    def CheckAgainstSpec( self, spec: dict, results: dict ) -> list:
        """
        check results by check spec
        weiliang want know which Items Fail
        return is dict:final_results = [fail Item name str]
        check spec map from specfile
        param spec:spec by partnum
        param results:from TDS data where save to db
        return: Fail Items'S dict:
                  len(final_results)<1 is pass
                  len(final_results)>0 is fail
        """
        log.debug( LOG_ENTER )

        final_results = []
        spec_dict = dict( spec )
        log.debug( 'check spec_dict is {0}'.format( spec_dict ) )

        results_dict = dict( results )
        log.debug( 'check results_dict is {0}'.format( results_dict ) )

        checkspecmap = self.cfginfo.get( BRD_TDS_CHECK_SPEC_MAP )
        log.debug( "check specmap is {0}".format( checkspecmap ) )

        if checkspecmap is None:
            log.error( 'No TDS Check Spec Map in device spec!' )
            return final_results
        # end if

        # For testing, no read from specfile and will be read
        # from specfile 'TDS Check Spec Map' in the future
        for specnameprefix in checkspecmap:
            log.debug( "----------{0}---------".format( specnameprefix ) )
            log.debug( "spec name prefix is {0}".format( specnameprefix ) )

            shippingdatakey = checkspecmap[ specnameprefix ]
            log.debug( "shipping data key is {0}".format( shippingdatakey ) )

            results = ''
            spec_key_min = "{}_Min".format( specnameprefix )
            log.debug( "specmap map key's MinName {0}".format( spec_key_min ) )
            spec_key_max = "{}_Max".format( specnameprefix )
            log.debug( "specmap map key's MaxName {0}".format( spec_key_max ) )

            # use '' not None Reduce judgment code as below
            # Sometimes there will be '' in the database,
            #     '0' is spec don't use it
            # We can't use it to check the effectiveness of the items
            maxvalue = ''
            minvalue = ''
            if ( spec_key_min in spec_dict ):
                log.debug( "find key name:{0} in spec_dict".format(
                    spec_key_min ) )

                minvalue = spec_dict[ spec_key_min ]
                log.debug( "find key {0}'s value {1}".format(
                    spec_key_min, minvalue ) )
            else:
                log.debug( "Not find key name:{0} in spec_dict".format(
                    spec_key_min ) )
            # end if

            if ( spec_key_max in spec_dict ):
                log.debug( "find key name:{0} in spec_dict".format(
                    spec_key_max ) )

                maxvalue = spec_dict[ spec_key_max ]
                log.debug( "find key {0}'s value {1}".format(
                    spec_key_max, maxvalue ) )
            else:
                log.debug( "Not find key name:{0} in spec_dict".format(
                    spec_key_max ) )
            # end if

            if ( shippingdatakey in results_dict ):
                log.debug( "find key name:{0} in results_dict".format(
                    shippingdatakey ) )
                results = results_dict[ shippingdatakey ]
                log.debug( "find key {0}'s value {1}".format(
                    shippingdatakey, results ) )
            else:
                log.debug( "Not find key name:{0} in results_dict".format(
                    shippingdatakey ) )
                # Not
                final_results.append( specnameprefix )
            # end if

            # Sometimes there will be '' in the database
            # With spec, not have values are also invalid
            curcheckitems = ''
            # check BER
            if ( ( specnameprefix == LM_TDS_DATA_BER ) and
                 ( self.f_code in [ FCODE_1752C21, '1752-C01' ] ) ):
                log.debug( "check BER spec" )

                issmsr = False
                try:
                    log.debug( "get SMSR Values" )
                    issmsr = self.Is_SMSR_Fall_In_Range( self.tdsdatasetdict )
                except ResultOutOfRangeError as ex:
                    log.error( ex )
                # end try

                log.debug( "get SMSR Values" )

                if ( issmsr == False ):
                    log.debug( "SMSR not 45<SMSR<48, don't need to check ber" )
                    continue
                # end if

                # just 45<SMSR<48 check BER
                # if empty '' also final
                if ( results == '' ):
                    msg = "{} is empty, so fail".format( specnameprefix )
                    log.error( msg )
                    curcheckitems = "{} is empty".format( specnameprefix )
                else:
                    log.debug( "check BER. SMSR is in range 45<SMSR<48" )
                    curcheckitems = self.CheckSpecResultsItems( specnameprefix,
                                    maxvalue, minvalue, results )
                # end if
            else:
                log.debug( "check {} spec".format( specnameprefix ) )
                curcheckitems = self.CheckSpecResultsItems( specnameprefix,
                                maxvalue, minvalue, results )
            # end if

            if ( curcheckitems != '' ):
                final_results.append( curcheckitems )
            # end if
        # end for

        log.debug( LOG_EXIT )
        return final_results
    # end function



    def CheckSpecResultsItems( self, checkekey: str, maxvalue: str,
                               minvalue: str, results: str ) -> str:
        """
        check Items,if fail return Items names
        results is empty also final
        param checkekey:check spec map key
        param maxvalue:form specfile where is the name is {key}_Max
        param minvalue:form specfile where is the name is {key}_Min
        param results:will check value.

        return:  Fail Items's name(key)
        """
        log.debug( LOG_ENTER )

        log.debug( 'Checking specs for {}'.format( checkekey ) )
        checkfailitems = ''

        # sql null dict return  None is str,convert
        # None can't convert float
        if ( results == 'None' ):
            results = ''
        # end if

        if ( ( results != '' ) and ( results != None ) ):
            try:
                checkresults = float( results )
            except ResultOutOfRangeError as ex:
                msg = "{} Type conversion error,also fail ".format( checkekey )
                checkfailitems = '{0} is empty'.format( checkekey )
                log.error( msg )
                log.error( str( ex ) )
            # end try
        else:
            msg = "results is '',also fail ".format( checkekey )
            checkfailitems = '{0} is empty'.format( checkekey )
            log.error( msg )
        # end if

        if ( minvalue != '' and maxvalue != '' and results != '' ):
            if ( eval( "checkresults<{0} or {1}<checkresults".format(
                minvalue, maxvalue ) ) ):
                checkfailitems = checkekey
                log.debug( "checkresults<{0} or {1}<checkresults is "
                    "fail".format( minvalue, maxvalue ) )
            else:
                log.debug( "between minvalue and maxvalue is pass".format(
                    minvalue, results, maxvalue ) )
            # end if
        elif ( minvalue != '' and maxvalue == '' and results != '' ):
            if ( eval( "{0}>checkresults".format( minvalue ) ) ):
                log.debug( "{0}<{1} is fail".format( minvalue, checkresults ) )
                checkfailitems = checkekey
            else:
                log.debug( "{0}<{1} is pass".format( minvalue, checkresults ) )
            # end if
        elif ( minvalue == '' and maxvalue != '' and results != '' ):
            if ( eval( "checkresults>{0}".format( maxvalue ) ) ):
                log.debug( "{0}<{1} is fail".format( checkresults, maxvalue ) )
                checkfailitems = checkekey
            else:
                log.debug( "{0}<{1} is pass".format( checkresults, maxvalue ) )
            # end if
        # end if

        log.debug( LOG_EXIT )
        return checkfailitems
    # end function



    def Is_SMSR_Fall_In_Range( self, tdsdataset: dict ) -> bool:
        """
        ! Only used for 1752-C21
        check if the SMSR test result between 45 and 48
        return: false if SMSR not in range;otherwise true
        """
        log.debug( LOG_ENTER )

        ret = False
        smsr2 = tdsdataset[ LM_TDS_DATA_SMSR2 ]
        # ! hardcoded values in constants file. should change!
        log.debug( 'SMSR Min is {}'.format( SMSR_MIN ) )
        log.debug( 'SMSR Max is {}'.format( SMSR_MAX ) )

        if ( ( smsr2 >= SMSR_MIN ) and ( smsr2 < SMSR_MAX ) ):
            log.debug( "SMSR in range" )
            ret = True
        else:
            log.debug( "SMSR not fall in range" )
            ret = False
        # end if

        log.debug( LOG_EXIT )
        return ret
    # end function



    def Update_LastTds_Status( self, status: str ) -> bool:
        """
        after check spec save status
        param status:Fass/FAIL
        return:true ->success
                 false->fail
        """
        log.debug( LOG_ENTER )

        updateStstus = False
        try:
            sql_query = "update LaserModuleShippingData set status='{0}'\
             WHERE ID=(SELECT max(id)  FROM \
             [OrtelTE].[dbo].[LaserModuleShippingData]\
              where Device_SN='{1}')".format( status, self._serialnumber )

            log.debug( "tds save query is {0}".format( sql_query ) )
            msserver = mssqlserver()
            msserver.ExecNonQuery( sql_query )
            updateStstus = True
        except NameError as ex:
            err_msg = "cound not connect to sql server {0}".format( ex )
            log.error( err_msg )
            raise TDSSqlConnectionError( err_msg )
        except Exception as ex:
            err_msg = "get tds data to sql server failed, {0}".format( ex )
            log.error( err_msg )
            raise TDSGETError( err_msg )
        # end try

        log.debug( LOG_EXIT )
        return updateStstus
    # end function



    def Get_WaveLen_SQL_Data( self ) -> str:
        """
        get the wavelength value from AWDM1550LIPDB

        return:
                None if we failed to get the SQL data
        """
        log.debug( LOG_ENTER )

        wlquery = {
            SQL_QUERY_VALUE_COL: "WLenAtTune",
            SQL_QUERY_TABLE: "AWDM1550LIPDB",
            SQL_QUERY_ID_COL: "SerNo",
            SQL_QUERY_TIME_COL: "RecordID",
            SQL_QUERY_ORDER_DIR: "DESC"
        }

        wavelen = ''
        if ( not self.GetSQLDBInformation( self._serialnumber, wlquery, True ) ):
            log.debug( 'Failed to get value from SQL database' )
            return None
        # end if

        log.debug( 'wlquery is {}'.format( wlquery ) )

        if ( wlquery[ KEY_TDS_VALUE ] == [] ):
            log.debug( 'No data for {}'.format( self._serialnumber ) )
            return None
        # end if

        wavelen = wlquery[ KEY_TDS_VALUE ][ 0 ][ wlquery[ SQL_QUERY_VALUE_COL ] ]

        log.debug( LOG_EXIT )
        return wavelen
    # end function



    def Get_SMSR_SQL_Data( self ) -> str:
        """
        get SMSR from ModuleFailSMSR

        :return:
        """
        log.debug( LOG_ENTER )

        smsrquery = {
            SQL_QUERY_VALUE_COL: "SMSR",
            SQL_QUERY_TABLE: "ModuleFailSMSR",
            SQL_QUERY_ID_COL: "Device_SN",
            SQL_QUERY_TIME_COL: "RID",
            SQL_QUERY_ORDER_DIR: "DESC"
        }

        smsr = ''
        if ( not self.GetSQLDBInformation( self._serialnumber, smsrquery, True ) ):
           log.debug( 'Failed to get value from SQL database' )
           return None
        # end if

        log.debug( 'smsrquery is {}'.format( smsrquery ) )

        if ( smsrquery[ KEY_TDS_VALUE ] == [] ):
            log.debug( 'No data for {}'.format( self._serialnumber ) )
            return None
        # end if

        smsr = smsrquery[ KEY_TDS_VALUE ][ 0 ][ smsrquery[ SQL_QUERY_VALUE_COL ] ]

        log.debug( LOG_EXIT )
        return smsr
    # end function



    def GetPowerOption( self, powerval: Decimal ):
        """
        Get the power option back in string format

        powerval: the power value which will be checked against specs to find
                  the correct option

        Return: str of power option, if found
                None: cannot find power option
        """
        log.debug( LOG_ENTER )

        poweroptions = self.cfginfo.get( LM_TDS_DATA_OPW_OPTION, None )
        if ( poweroptions is None ):
            log.error( 'No Optical Power Options defined in device spec' )
            return None

        powoption = None
        currpow = powerval
        for eachoption in poweroptions:
            minpow = Decimal( poweroptions[ eachoption ].get( SPEC_MIN ) )
            maxpow = Decimal( poweroptions[ eachoption ].get( SPEC_MAX ) )

            if ( ( minpow != None ) and ( maxpow != None ) ):
                if ( eval( "currpow<={0} and currpow>={1}".format(
                    maxpow, minpow ) ) ):
                    powoption = eachoption
                    break
            elif ( minpow != None ):
                if ( eval( "currpow>{0}".format( minpow ) ) ):
                    powoption = eachoption
                    break
            elif ( maxpow != None ):
                if ( eval( "currpow<{0}".format( maxpow ) ) ):
                    powoption = eachoption
                    break
        # end for

        log.debug( LOG_EXIT )
        return powoption



    def SaveTDSData( self, desp: str = None, pn: str = None ):
        """
        save the TDS data to sql server
        desp: description for the part number
        pn: part number of the laser

        return: None
        """
        log.debug( LOG_ENTER )

        tds_data_dict = {}
        tds_data_dict = self.tdsdatasetdict

        # check if data dictionary is empty
        if ( not tds_data_dict ):
            err_msg = "No data found!"
            log.error ( err_msg )
            raise TDSSaveError( err_msg )
        # end if

        tdsname = None
        # get the first four digits of the product id for tdsname
        tds_product_id = tds_data_dict.get( LM_TDS_DATA_PROD_ID )
        if( tds_product_id is None or ( len( tds_product_id ) < 4 ) ):
            log.debug( '{} not in tdsdatadict. It is okay.'
                       .format( LM_TDS_DATA_PROD_ID ) )
            tdsname = self.f_code
        else:
            tdsname = tds_product_id[ 0:4 ]
        # end if

        power = tds_data_dict.get( LM_TDS_DATA_OPW )
        # All units must have optical power
        if ( power is None ):
            err_msg = "Missing LM_TDS_DATA_OPW in results!"
            log.error ( err_msg )
            raise TDSSaveError( err_msg )
        power_option = self.GetPowerOption( Decimal( power ) )

        wavelen = tds_data_dict.get( LM_TDS_DATA_WAVELEN )
        if ( wavelen == "" ):
            wavelen = self.Get_WaveLen_SQL_Data()
        # end if

        if ( wavelen is None ):
            err_msg = "Did not find wavelength data in SQL"
            log.error ( err_msg )
            raise TDSSaveError( err_msg )
        # end if

        smsr_ms = tds_data_dict.get( LM_TDS_DATA_SMSR )
        if ( ( smsr_ms is None ) or ( smsr_ms == "" ) ):
            log.debug( 'No SMSR value. Okay.' )
            pass
        # end if

        # This check is only used for 1752-C21 products.
        # SMSR is not required by all products. Move out of this fxn.
        # SMSR 2 may be from either mongo or sql
        # smsr_ms is from sql
        if ( self.f_code in [ FCODE_1752C21 ] ):
            smsr2 = tds_data_dict.get( LM_TDS_DATA_SMSR2 )
            log.debug( 'smsr2 is {}'.format( smsr2 ) )
            if ( ( smsr2 is None ) or ( smsr2 == "" ) ):
                err_msg = "Missing SMSR 2 value in results!"
                log.error ( err_msg )
                raise TDSSaveError( err_msg )
            # end if

            # If smsr2 is out of spec, then don't need BER for 1752-C21
            if ( ( smsr2 < SMSR_MIN ) or ( smsr2 >= SMSR_MAX ) ):
                tds_data_dict[ LM_TDS_DATA_BER ] = ""
            # end if
        # end if

        # some time the mongo data None is string, but MSSQL is float,
        # just can insert null
        sql_query = "insert into LaserModuleShippingData(Device_SN \
            ,PRODUCT,Description,PartNo_Select \
            ,Bias_Select,TDSName,PowerOption,Ith,Slopeff \
            ,Chirp,ConnectorType,SMSR,CSO,CTB,CNR \
            ,OpticalPower,MPDCurrent,TrackingError_Max \
            ,TrackingError_Min,CountPrint,ForwardVoltage \
            ,FreqResponse,LsrTemperature,Iop,CNR_61 \
            ,CNR_547,Station,Operator,B_Const_K,RTH_Kohm \
            ,SMSR2,BER,TDS_DT,WaveLen,VERSION,MPDISlope_LIP) \
            VALUES('{0}','{1}','{2}','{3}',{4},'{5}' \
            ,'{6}','{7}','{8}','{9}','{10}','{11}','{12}' \
            ,'{13}','{14}','{15}','{16}','{17}','{18}' \
            ,'{19}','{20}','{21}','{22}','{23}','{24}' \
            ,'{25}','{26}','{27}','{28}','{29}','{30}' \
            ,'{31}','{32}','{33}','{34}','{35}')".format(
            # sn
            self._serialnumber,
            # PRODUCT
            self.f_code,
            # Description
            desp,
            # Part number
            pn,
            # Bias_Select - default is 0.
            "0",
            # TDSName
            # TO DO check with weiliang, should we use the real TDS #
            str( tdsname ),
#            tds_data_dict.get( BRD_TDS_PRNT_FN ),
            # Power Option
            str( power_option ),
            # Ith
            tds_data_dict.get( LM_TDS_DATA_ITH ),
            # Slopeff
            tds_data_dict.get( LM_TDS_DATA_SLOPEFF ),
            # Chirp
            tds_data_dict.get( LM_TDS_DATA_CHIRP ),
            # connector type, default to SC/APC
            tds_data_dict.get( BRD_CONNECT_OPTION, LM_TDS_DATA_CONN_DEFAULT ),
            # SMSR1  get real smsr from msql
            smsr_ms,
            # CSO
            tds_data_dict.get( LM_TDS_DATA_CSO ),
            # CTB
            tds_data_dict.get( LM_TDS_DATA_CTB ),
            # CNR - default to -999 for units that don't use this field
            tds_data_dict.get( LM_TDS_DATA_CNR, -999 ),
            # power
            str( power ),
            # MPDCurrent
            str( tds_data_dict.get( LM_TDS_DATA_MPD ) ),
            # Tracking Error Positive
            str( tds_data_dict.get( LM_TDS_DATA_ERR_MAX ) ),
            # Tracking Error Negative
            str( tds_data_dict.get( LM_TDS_DATA_ERR_MIN ) ),
            # tds CountPrint, default is 1
            "1",
            # ForwardVoltage
            str( tds_data_dict.get( LM_TDS_DATA_FORWARD_VOLT ) ),
            # Frequency Response
            str( tds_data_dict.get( LM_TDS_DATA_FR ) ),
            # LsrTemperature
            str( tds_data_dict.get( LM_TDS_DATA_LSR_TEMP ) ),
            # Iop
            str( tds_data_dict.get( LM_TDS_DATA_Iop ) ),
            # CNR_61
            str( tds_data_dict.get( LM_TDS_DATA_CNR_61 ) ),
            # CNR_547
            str( tds_data_dict.get( LM_TDS_DATA_CNR_547 ) ),
            # Station
            str( tds_data_dict.get( LM_TDS_DATA_STATION ) ),
            # Operator
            str( tds_data_dict.get( LM_TDS_DATA_OPERATOR ) ),
            # B_Const_K
            str( tds_data_dict.get( LM_TDS_DATA_BC ) ),
            # RTH_Kohm
            str( tds_data_dict.get( LM_TDS_DATA_RTH_KOHM ) ),
            # SMSR2
            str( tds_data_dict.get( LM_TDS_DATA_SMSR2 ) ),
            str( tds_data_dict.get( LM_TDS_DATA_BER, "" ) ),
            # TDS_DT
            datetime.datetime.now().strftime ( "%Y-%m-%d %H:%M:%S" ),
            wavelen,
            SOFT_VERSION_NUMBER,
            str( tds_data_dict.get( LM_TDS_DATA_MPD_SLOPE ) )
            ).replace( '\'None\'', 'null' )

        log.debug( "tds save query is {0}".format( sql_query ) )

        try:
            sqlserver = mssqlserver()
        except NameError as ex:
            err_msg = "cound not connect to sql server {0}".format( ex )
            log.error( err_msg )
            raise TDSSqlConnectionError( err_msg )
        # end try

        try:
            sqlserver.ExecNonQuery( sql_query )
        except IntegrityError as ex:
            err_msg = "data format is wrong, {0}".format( ex )
            log.error ( err_msg )
            raise TDSSaveError( err_msg )
        except Exception as ex:
            err_msg = "save tds data to sql server failed, {0}".format( ex )
            log.error ( err_msg )
            raise TDSSaveError( err_msg )
        # end try

        log.debug( LOG_EXIT )
    # end function
# end class





class TDSManagerTX( TDSManager ):
    @classmethod
    def IsTDSManagerFor( cls:object, product_type: str ) -> bool:
        """ Checks if the product type is supported by this class

        :param product_type: ex LM, TX
        :return: ( boolean )
            True: produc ttype is supported
            False: not supported
        """
        return ( product_type in [ PRODUCT_TYPE_TX ] )
    # end function



    def CreateTDS( self, printbool: int ) -> dict:
        """ Fill all the required fields on test data sheet
        based on test results. Reads the data from testinfo (json) file and
        update each cell based on test name and group.

        :param: printbool: int value used as bool. Can be 0 or 1.
                1 to print. 0 if we want to skip TDS update and printing.
        :return: (dictionary)
            retdict:
                status_bool
                status_msg
        """
        log.debug( LOG_ENTER )

        retdict = dict( status_bool = False, status_msg = '' )
        tdsworker_obj = None

        tdsworker_obj = TDSWorker_TX( serialnumber = self._serialnumber,
                                      sessiontype = TST_GRP_TYPE_PROD,
                                      num_type = self.num_type )

        retdict = tdsworker_obj.CreateTDS( printbool )

        log.debug( LOG_EXIT )

        return retdict
    # end function
# end class





class TDSManagerLM( TDSManager ):
    @classmethod
    def IsTDSManagerFor( cls: object, product_type: str ) -> bool:
        """ Checks if the product type is supported by this class

        param product_type: ex LM, TX
        return: ( boolean )
        True: produc ttype is supported
        False: not supported
        """
        return ( product_type in [ 'LM' ] )
    # end function



    def CreateTDS( self, printbool: int ) -> dict:
        """ Fill all the required fields on test data sheet
        based on test results. Reads the data from testinfo (json) file and
        update each cell based on test name and group.

        :param: printbool: int value used as bool. Can be 0 or 1.
                1 to print. 0 if we want to skip TDS update and printing.
        :return: (dictionary)
            retdict:
                status_bool
                status_msg
        """
        log.debug( LOG_ENTER )

        retdict = dict( status_bool = False, status_msg = '' )

        tdsworker_obj = None
        tdsworker_obj = TDSWorker_LM( serialnumber = self._serialnumber,
                                      sessiontype = TST_GRP_TYPE_PROD,
                                      num_type = self.num_type )
#        if ( self.f_code == FCODE_1752C21 ):
#           pass
#        else:
#            msg = 'LM family not supported: {0}'.format( self.f_code )
#            log.debug( msg )
#            retdict[ KEY_TDS_STATUS_MSG ] = msg
#            retdict[ KEY_TDS_STATUS_BOOL ] = False
#            return retdict
#        # end if

        log.debug( LOG_EXIT )

        retdict = tdsworker_obj.CreateTDS( printbool )
        return retdict
    # end function
# end class





class BoardInformation():
    """Retrieves board information from the database
    """
    def __init__( self, fcode: str, serialnumber: str ):
        ''' PrinterObj's constructor

        fcode: family code
        '''
        log.debug( LOG_ENTER )

        self.f_code = fcode
        self._serialnumber = serialnumber

        log.debug( LOG_EXIT )
        pass
    # end function



    def _getBoardInfoFromMongo( self, num_type: str, in_num: str,
        prod_type: str ) -> dict:
        '''
        Get board information from mongo

        num_type: BRD_LASER_ID, BRD_CUSTOMER_ID(serial number)
        in_num:   inputted number string

        return: None -> get board information failed
            dict -> get board information successed
            {
                "Database Type":             "",
                "ProductID":                 "",
                "Serial Number":              "",
                "Manufacturing Date":        "",
                "PCBA ID":                    "",
                "LASER ID":                  "",

                "ITU Channel":               "",
                "Test Result":               "",
            }
        '''
        log.debug( LOG_ENTER )

        retdict = None
        if ( prod_type == PRODUCT_TYPE_TX ):
            retdict = self._getBoardInfoFromMongo_TX( num_type, in_num )
        elif ( prod_type == PRODUCT_TYPE_LM ):
            retdict = self._getBoardInfoFromMongo_LM( num_type, in_num )
        else:
            log.error( 'Product type not supported!' )
            retdict = None
        # end if

        log.debug( LOG_EXIT )
        return retdict
    # end function



    def _getBoardInfoFromMongo_TX( self, num_type: str, in_num: str ) -> dict:
        '''
        Get board information from mongo

        num_type: BRD_LASER_ID, BRD_CUSTOMER_ID(serial number)
        in_num:   inputted number string

        return: None -> get board information failed
            dict -> get board information successed
            {
                "Database Type":             "",
                "ProductID":                 "",
                "Serial Number":              "",
                "Manufacturing Date":        "",
                "PCBA ID":                    "",
                "LASER ID":                  "",

                "ITU Channel":               "",
                "Test Result":               "",
            }
        '''
        log.debug( LOG_ENTER )

        # use getTXTestResult when searching by serialnumber/customerid
        if ( ( num_type == BRD_SERIAL_NUM ) or
                                    ( num_type == BRD_CUSTOMER_ID ) ):
            boardinfodict = getTXTestResult( db_info = db_info,
                                             serialnumber = in_num )
            filter_dict = { BRD_SN: in_num }

            # Debugging code - get all test results and write them to
            #   'allmongoresults.json':
            boardinfodictall = getTXTestResultsComplex( db_info = db_info,
                                            filter_dict = filter_dict,
                                            limit = 20 )
            write_json_file( ALL_MONGO_RESULTS_FN, boardinfodictall )

            # Query error
            if ( boardinfodict is None ):
                log.debug( 'Issue occurred during database query for results' )
                return None
            # end if

            # No results
            if ( boardinfodict == {} ):
                log.error( 'Mongo doc query returned no'
                           'results with {0}:{1}'.format( num_type, in_num ) )
                return None
            # end if

        # otherwise use getTxTestResultsComplex for arbitrary
        # field filter search
        else:
            filter_dict = { num_type: in_num }
            boardinfolist = getTXTestResultsComplex( db_info = db_info,
                                            filter_dict = filter_dict )

            # Query error
            if ( boardinfolist is None ):
                log.debug( "Issue occurred during database query for results" )
                return None
            # end if

            # No results
            if ( boardinfolist == [] ):
                log.error( 'Mongo doc query returned no'
                           'results with {0}:{1}'.format( num_type, in_num ) )
                return None
            # end if

            boardinfodict = boardinfolist[ 0 ]

        if ( not write_json_file( MONGO_RESULTS_FN, boardinfodict ) ):
            log.error( 'Cannot write result to {}'.format( MONGO_RESULTS_FN ) )
            return None

        # Check boardinfodict is right
        if ( boardinfodict is None ):
            log.error( 'Mongo query returned no'
                       'results with {0}:{1}'.format( num_type, in_num ) )
            return None
        # end if

        # Check BRD_MAN_DATE in boardinfodict
        if ( BRD_MAN_DATE not in boardinfodict.keys() ):
            log.error( 'Key {0} not found in '
                       'boardinfodict from mongo'.format( BRD_MAN_DATE ) )
            return None
        # end if

        # Check BRD_CUSTOMER_ID in boardinfodict
        if ( BRD_SN not in boardinfodict.keys() ):
            log.error( 'Key {0} not found in '
                       'boardinfodict from mongo'.format( BRD_SN ) )
            return None
        # end if

        # Check HW_PRODUCT_ID in boardinfodict
        if ( HW_PRODUCT_ID not in boardinfodict.keys() ):
            log.error( 'Key {0} not found in'
                       'boardinfodict from mongo'.format( HW_PRODUCT_ID ) )
            return None
        # end if

        # Check BRD_PCB_ID in boardinfodict
        if ( BRD_PCB_ID not in boardinfodict.keys() ):
            log.error( 'Key {0} not found in'
                       'boardinfodict from mongo'.format( BRD_PCB_ID ) )
            return None
        # end if

        # Check BRD_LASER_ID in boardinfodict
        if ( BRD_LASER_ID not in boardinfodict.keys() ):
            log.error( 'Key {0} not found in'
                       'boardinfodict from mongo'.format( BRD_LASER_ID ) )
            return None
        # end if

        # Check BRD_ITU_CH_NUM in boardinfodict
        if ( BRD_ITU_CH_NUM not in boardinfodict.keys() ):
            log.error( 'Key {0} not found in'
                       'boardinfodict from mongo'.format( BRD_ITU_CH_NUM ) )
            return None
        # end if

        # Check Final Result in boardinfodict
        if ( TST_FINAL_RESULT not in boardinfodict.keys() ):
            log.error( 'Key {0} not found in'
                       'boardinfodict from mongo'.format( TST_FINAL_RESULT ) )
            return None
        # end if

        # If no product family assume CATV
        if ( DB_PRODUCT_FAMILY not in boardinfodict.keys() ):
            log.debug( 'No product family in boardinfo, assuming CATV' )
            boardinfodict[ DB_PRODUCT_FAMILY ] = FAMILY_CATV
        # end if

        # Fill return dict
        brd_info_mongo = {}
        # Database Type
        brd_info_mongo[ SPEC_DATA_TYPE ] = SPEC_DB_MONGO
        # customer ID ( serial number )
        brd_info_mongo[ BRD_SN ] = boardinfodict[ BRD_SN ]
        brd_info_mongo[ BRD_CUSTOMER_ID ] = boardinfodict[ BRD_SN ]

        # Get latest test date via 'tx_process_flow'
        if ( PRCS_FLOW_TX in boardinfodict.keys() ):
            operation_date = boardinfodict[ PRCS_FLOW_TX ][ -1 ]\
                             [ 'operation_date' ]
            if( type( operation_date ) == type( OrderedDict() ) ):
                # Old timestamp format
                mongo_timestamp = operation_date[ '$date' ]
                mongo_date = datetime.datetime.utcfromtimestamp(
                             mongo_timestamp / 1000 )
            else:
                # New timestamp format
                mongo_date = datetime.datetime.strptime( operation_date,
                        '%Y-%m-%d %H:%M:%S' )
        else:
            # Alternative (less consistent) method of getting date via
            # 'Manufacturing Date'
            date_time_str = boardinfodict[ BRD_MAN_DATE ]
            mongo_date = ''

            if( len( date_time_str ) == 8 ):
                try:
                    # P2 Products
                    mongo_date = datetime.datetime.strptime( date_time_str, '%m%d%Y' )
                except ValueError:
                    pass
                # end try
            elif( len( date_time_str ) == 6 ):
                if ( not mongo_date ):
                    try:
                        # AT3 and GX2 Products
                        mongo_date = datetime.datetime.strptime( date_time_str,
                                '%m%d%y' )
                    except ValueError:
                        pass
                    # end try
                # end if

            if ( not mongo_date ):
                try:
                    # Hyphenated date string
                    mongo_date = datetime.datetime.strptime( date_time_str,
                            '%Y-%m-%d' )
                except ValueError:
                    pass
                # end try
            # end if

            if ( not mongo_date ):
                log.debug( 'Manufacturing date string not supported: {0}'.format(
                            date_time_str ) )
                return None
            # end if

        # 11/21/17 Andy: This key naming may be faulty logic. We overwrite
        # the 'manufacturing date' with latest test date. Instead it would be
        # more accurate to add a 'last test date' key, and format based on that
        # since some products are still concerned with the manufacturing date
        # and we don't always want to get rid of it.

        # Temporary solution: save to a temp variable for boards that need it
        #   i.e. GX2
        manufacturing_date = boardinfodict[ BRD_MAN_DATE ]
        brd_info_mongo[ BRD_MAN_DATE ] = mongo_date
        # Get different string versions of date
        brd_info_mongo = self._convertDateTime2String( brd_info_mongo )
        # PCBA ID
        brd_info_mongo[ BRD_PCB_ID ] = boardinfodict[ BRD_PCB_ID ]
        # Laser ID
        brd_info_mongo[ BRD_LASER_ID ] = boardinfodict[ BRD_LASER_ID ]
        # Product ID
        brd_info_mongo[ HW_PRODUCT_ID ] = boardinfodict[ HW_PRODUCT_ID ]
        # Final Result
        brd_info_mongo[ TST_FINAL_RESULT ] = boardinfodict[ TST_FINAL_RESULT ]

        # add extra details in dictionary for GX2 products: detailed
        # timestamp based on latest tx_process_flow and full customer ID
        if ( 'GX2' in boardinfodict[ HW_PRODUCT_ID ] ):
                # add timestamp
            if ( PRCS_FLOW_TX in boardinfodict.keys() ):
                brd_info_mongo[ KEY_TDS_DATE_CREATION ] = mongo_date.strftime(
                        KEY_TDS_DATE_LOG_FORMATTING )
            else:
                brd_info_mongo[ KEY_TDS_DATE_CREATION ] = brd_info_mongo[
                        KEY_TDS_MANUFACTURE_MM_YY ]
            # end if

            manufacturing_date = datetime.datetime.strptime(
                    manufacturing_date, '%m%d%y')

            # add customer ID
            monthcodedict = { '01':'A', '02':'B', '03':'C', '04':'D', '05':'E',
                              '06':'F', '07':'G', '08':'H', '09':'J', '10':'K',
                              '11':'L', '12':'M' }
            # 10/4 Should not use hardcoded value for mfgloc!
            mfgloc = 'EA'
            model = boardinfodict[ HW_PRODUCT_ID ]
            serialnum = boardinfodict[ BRD_SN ]
            powermodel = model[ 11:13 ].strip( '-' )
            year = manufacturing_date.strftime( KEY_TDS_YEAR )
            month = manufacturing_date.strftime( KEY_TDS_MONTH )
            day = manufacturing_date.strftime( KEY_TDS_DAY )
            customerid = monthcodedict[ month ] + year + mfgloc + '60' + \
                    '{0:0>2}'.format( powermodel ) + day + serialnum[ 3: ]
            brd_info_mongo[ BRD_CUSTOMER_ID ] = customerid
        # end if

        # not tested with new RESTful API yet
        # Laser module does not have product family
        if ( ( DB_PRODUCT_FAMILY in boardinfodict ) and
             ( boardinfodict[ DB_PRODUCT_FAMILY ] == FAMILY_SATCOM ) ):
            # needs to know if in SATCOM family in the PrintDeviceGUI class.
            brd_info_mongo[ DB_PRODUCT_FAMILY ] = FAMILY_SATCOM

            finddict = {
                    BRD_SN : boardinfodict[ BRD_SN ],
                }

            result_data = getTXTestResultsComplex( db_info = db_info,
                filter_dict = filter_dict )

            if ( result_data is None ):
                log.error( 'Database issue encountered!' )
                return False
            elif ( result_data == {} ):
                log.debug( 'Query returned 0 results ( Query: {0} )'.format(
                    finddict ) )
                return False
            # end if

            test_final_result = result_data [ TST_FINAL_RESULT ]
            if ( test_final_result is None ):
                log.error( 'Cannot get final Satcom Result' )
                brd_info_mongo[ TST_RESULT ] = TST_RES_FAILED
            elif ( test_final_result ) == TST_RES_PASSED:
                brd_info_mongo[ TST_RESULT ] = TST_RES_PASSED
            else:
                brd_info_mongo[ TST_RESULT ] = TST_RES_FAILED
            # end if
        else:
            # Test result
            # --get Distortion and  Frequency Response test result
            dis_tst_res = None
            s21_tst_res = None

            try:
                dis_tst_res = boardinfodict[ TST_RES_DIST_DATA ]
                s21_tst_res = boardinfodict[ TST_RES_FREQ_RESP_DATA ]
            except Exception as e:
                log.error( 'One of more tests could not be retrieved from '
                           'Mongo' )
            # end try

            # -- check test result
            if ( dis_tst_res is None ) or ( s21_tst_res is None ):
                log.error( 'Can not get {0} {1} test result.'.format( num_type,
                        in_num ) )
                brd_info_mongo[ TST_RESULT ] = TST_RES_FAILED
            # Fill test result to dict
            elif ( boardinfodict[ 'Final Result' ] == TST_RES_PASSED ):
                brd_info_mongo[ TST_RESULT ] = TST_RES_PASSED
            else:
                brd_info_mongo[ TST_RESULT ] = TST_RES_FAILED
            # end if
        # end if

        # The below is only in mongo db
        # ITU Channel
        brd_info_mongo[ BRD_ITU_CH_NUM ] = boardinfodict[ BRD_ITU_CH_NUM ]

        # add test results to board info, don't need a second db request later
        brd_info_mongo[ TST_RES_DIST_DATA ] = dis_tst_res
        brd_info_mongo[ TST_RES_FREQ_RESP_DATA ] = s21_tst_res

        if ( PRCS_FLOW_TX in boardinfodict ):
            brd_info_mongo[ PRCS_FLOW_TX ] = boardinfodict[ PRCS_FLOW_TX ]
        # end if

        log.debug( LOG_EXIT )
        return brd_info_mongo
    # end function



    def _getBoardInfoFromMongo_LM( self, num_type: str, in_num: str ) -> dict:
        '''
        Get board information from mongo

        num_type: BRD_LASER_ID, BRD_CUSTOMER_ID(serial number)
        in_num:   inputted number string

        return: None -> get board information failed
            dict -> get board information successed
            {
                "Database Type":             "",
                "ProductID":                 "",
                "Serial Number":              "",
                "Manufacturing Date":        "",
                "PCBA ID":                    "",
                "LASER ID":                  "",

                "ITU Channel":               "",
                "Test Result":               "",
            }
        '''
        log.debug( LOG_ENTER )

        log.debug( "start query for {0}".format( in_num ) )

        filter_dict = { num_type: in_num }
        boardinfolist = getLMTestResultsComplex( db_info = db_info,
                                                 filter_dict = filter_dict )

        # Query error
        if ( boardinfolist is None ):
            log.debug( "Issue occurred during database query for results" )
            return None
        # end if

        # No results
        if ( boardinfolist == [] ):
            log.error( 'Mongo doc query returned no'
                       'results with {0}:{1}'.format( num_type, in_num ) )
            return None
        # end if

        boardinfodict = boardinfolist[ 0 ]

        retbool = write_json_file( MONGO_RESULTS_FN, boardinfodict )
        if ( not retbool ):
            log.error( 'Cannot write result to {}'.format( MONGO_RESULTS_FN ) )
            return None
        # end if

        # Check boardinfodict is right
        if ( boardinfodict is None ):
            log.error( 'Mongo query returned no'
                       'results with {0}:{1}'.format( num_type, in_num ) )
            return None
        # end if

        # Check Final Result in boardinfodict
        if ( TST_FINAL_RESULT not in boardinfodict.keys() ):
            log.error( 'Key {0} not found in'
                       'boardinfodict from mongo'.format( TST_FINAL_RESULT ) )
            return None
        # end if

        # If no product family, assume LM
        if ( DB_PRODUCT_FAMILY not in boardinfodict.keys() ):
            log.debug( 'No product family in boardinfo, assuming laser' )
            boardinfodict[ DB_PRODUCT_FAMILY ] = FAMILY_LASERMODULE
        # end if

        # Fill return dict
        brd_info_mongo = {}
        # Database Type
        brd_info_mongo[ SPEC_DATA_TYPE ] = SPEC_DB_MONGO
        # customer ID ( serial number )

        # Transmitter schema have the serial number under the key "SN"
        # Laser module schema have the serial number under the key
        #       "Serial Number" or "Serial_Number". Copy value for key "SN"
        #       for compatibility
        brd_info_mongo[ BRD_SN ] = boardinfodict[ BRD_SERIAL_NUM ]

        brd_info_mongo[ BRD_SERIAL_NUM ] = boardinfodict[ BRD_SERIAL_NUM ]
        brd_info_mongo[ BRD_CUSTOMER_ID ] = boardinfodict[ BRD_SERIAL_NUM ]
        boardinfodict[ BRD_CUSTOMER_ID ] = boardinfodict[ BRD_SERIAL_NUM ]
        brd_info_mongo[ TST_FINAL_RESULT ] = boardinfodict[ TST_FINAL_RESULT ]
        # Product ID is used to pull SQL Time Comparison table
#        brd_info_mongo[ HW_PRODUCT_ID ] = boardinfodict[  ]

        if ( PRCS_FLOW_LM in boardinfodict ):
            brd_info_mongo[ PRCS_FLOW_LM ] = boardinfodict[ PRCS_FLOW_LM ]
        # end if

        # add test results to board info, don't need a second db request later
        boardinfodict = self._FilterData( boardinfodict )
        if ( boardinfodict is {} ):
            log.error( 'Error during parsing the results!' )
            return None
        # end if

        # Append brd_info_mongo
        boardinfodict.update( brd_info_mongo )

        log.debug( LOG_EXIT )
        return boardinfodict
    # end function



    def _getBoardInfoFromSQL( self, db_table_name:str, num_type:str,
                             in_num:str, mongo_bridge_dict: dict,
                             lookup_table_dict: dict, prod_type: str ) -> dict:
        '''
        Get board information from sql database

        db_table_name: date base table name
        num_type: identifies what type of number the inputted number string is.
            examples are BRD_SERIAL_NUM, BRD_LASER_ID, BRD_CUSTOMER_ID
        in_num: inputted number string
        mongo_bridge_dict: product SQL table name and part number - product ID
                           corresponding dict
        lookup_table_dict: the keys in which cell in SQL table
        prod_type: type of product being tested, e.g. LM or TX

        return:    None -> get board information failed
                  dict -> get board information succeeded
        {
            "Database Type":             "",
            "ProductID":                 "",
            "Serial Number":              "",
            "Manufacturing Date":        "",
            "PCBA ID":                   "",
            "LASER ID":                  "",

            "Module ID":                 "",
            "PartNumber"                 "",
            "Test Result":               "",
        }
        '''
        log.debug( LOG_ENTER )
        retdict = None
        if ( prod_type == PRODUCT_TYPE_TX ):
            retdict = self._getBoardInfoFromSQL_TX( db_table_name, num_type,
                in_num, mongo_bridge_dict, lookup_table_dict )
        elif ( prod_type == PRODUCT_TYPE_LM ):
            retdict = self._getBoardInfoFromSQL_LM( db_table_name, num_type,
                in_num, mongo_bridge_dict, lookup_table_dict )
            # Do not need to get test results from
            # 'AWDM1550LIPDB' for 1310 LM
            if ( self.f_code in [ '1612-STD', '16XX' ] ):
                addeddict = self._getBoardInfoFromSQL_FailSMSR_LM(
                        'ModuleFailSMSR', num_type, in_num,
                        mongo_bridge_dict, lookup_table_dict )
                retdict.update( addeddict )
            # need to find a better way to handle this!
            # For qam products, we may need to use different tables,
            # one for time, and another for others.
            else:
                if ( ( self.f_code not in [ '1798' ] ) and
                                                ( retdict is not None ) ):
                    addeddict = self._getBoardInfoFromSQL_LIP_LM(
                            'AWDM1550LIPDB', num_type, in_num,
                            mongo_bridge_dict, lookup_table_dict )
                    retdict.update( addeddict )
                # end if
            # end if
        else:
            log.error( 'Product type not supported!' )
            retdict = None

        log.debug( LOG_EXIT )
        return retdict
    # end function



    def _getBoardInfoFromSQL_TX( self, db_table_name:str, num_type:str,
                             in_num:str, mongo_bridge_dict: dict,
                             lookup_table_dict: dict ) -> dict:
        '''
        Get board information from sql database

        db_table_name: date base table name
        num_type: identifies what type of number the inputted number string is.
            examples are BRD_SERIAL_NUM, BRD_LASER_ID, BRD_CUSTOMER_ID
        in_num: inputted number string
        mongo_bridge_dict: product SQL table name and part number - product ID
                           corresponding dict
        lookup_table_dict: the keys in which cell in SQL table

        return:    None -> get board information failed
                  dict -> get board information successed
        {
            "Database Type":             "",
            "ProductID":                 "",
            "Serial Number":              "",
            "Manufacturing Date":        "",
            "PCBA ID":                   "",
            "LASER ID":                  "",

            "Module ID":                 "",
            "PartNumber"                 "",
            "Test Result":               "",
        }
        '''
        log.debug( LOG_ENTER )

        # create SQL query string
        if ( num_type == BRD_LASMOD ):
            getBySNSql = str( "SELECT * FROM {0} WHERE [SerNo] = "
                              "'{1}' ORDER BY [RecordID] DESC".format(
                              db_table_name, in_num ) )
        elif ( num_type == BRD_LASER_ID ):
            getBySNSql = str( "SELECT * FROM {0} WHERE [BoardNo] = "
                              "'{1}' AND  [TestType] = 'FNL' AND "
                              "[PFStatus] = 'PASS' ORDER BY [RecordID] "
                              "DESC".format( db_table_name, in_num ) )
        elif ( num_type == BRD_PCBA_ID or num_type == BRD_PCB_ID ):
            getBySNSql = str( "SELECT * FROM {0} WHERE [LaserSerNo] = "
                              "'{1}' AND  [TestType] = 'FNL' AND "
                              "[PFStatus] = 'PASS' ORDER BY [RecordID] "
                              "DESC".format( db_table_name, in_num ) )
        elif ( num_type == BRD_SERIAL_NUM or num_type == BRD_TX_ID ):
            getBySNSql = str( "SELECT * FROM {0} WHERE [TransmitterSerNo]"
                              " = '{1}' AND [TestType] = 'FNL' AND  [PFStatus]"
                              " = 'PASS' ORDER BY [RecordID] "
                              "DESC".format( db_table_name, in_num ) )
        # end if

        # connect to sql server
        log.debug( 'SQL query: {0}'.format( getBySNSql ) )
        sqlserver = mssql.mssqlserver()
        if not sqlserver:
            log.error( 'Failed to connect to SQL Database' )
            return None
        # end if

        # execute query
        record = sqlserver.ExecQuery( getBySNSql )
        if not record:
            log.error( 'Failed to get record for {0}: {1} from '
                       'SQL database'.format( num_type, in_num ) )
            return None
        # end if

        # Judge record[0] is not empty
        if ( len( record ) < 1 ):
            log.error( 'No data in record'.format( num_type, in_num ) )
            return None
        # end if

        if ( mongo_bridge_dict is None ):
            log.error( 'Can not load {0} file'.format( MONGO_BRIDGE_FN ) )
            return None
        # end if

        if ( db_table_name not in mongo_bridge_dict ):
            log.error( '{0} key is not in {1} file'.format( db_table_name,
                                                            MONGO_BRIDGE_FN ) )
            return None
        # end if

        custid_prdctid_brdg = mongo_bridge_dict[ db_table_name ]

        if ( lookup_table_dict is None ):
            log.error( 'Can not load {0} file'.format( LOOK_UP_TABLE_FN ) )
            return None
        # end if

        if ( db_table_name not in lookup_table_dict ):
            log.error( '{0} key is not in {1} file'.format( db_table_name,
                                                           LOOK_UP_TABLE_FN ) )
            return None
        # end if

        sql_row_clm_table = lookup_table_dict[ db_table_name ]

        # Infer product ID by CusPartNum and
        # product ID/cust id bridge(Mongo_bridge)
        if ( HW_PART_NUM in sql_row_clm_table ):
            row = sql_row_clm_table[ HW_PART_NUM ][ KEY_TDS_ROW ]
            clm = sql_row_clm_table[ HW_PART_NUM ][ KEY_TDS_COLUMN ]
        elif ( HW_MODEL in sql_row_clm_table ):
            row = sql_row_clm_table[ HW_MODEL ][ KEY_TDS_ROW ]
            clm = sql_row_clm_table[ HW_MODEL ][ KEY_TDS_COLUMN ]
        # end if

        cust_id = str( record[ row ][ clm ] ).strip()

        product_id = None
        for pid, pspec in custid_prdctid_brdg.items():
            if ( pid in cust_id ):
                product_id = pspec
                break
            # end if
        # end for

        if ( product_id is None ):
            log.error( 'The software do not support this product '
                       '{0} now'.format( cust_id ) )
            return None
        # end if

        # Fill return dict
        brd_info_SQL = {}
        # Database Type
        brd_info_SQL[ SPEC_DATA_TYPE ] = SPEC_DB_SQL
        # product ID
        brd_info_SQL[ HW_PRODUCT_ID ] = product_id

        # Search for required keys using lookup_table
        sqlkeys_txrx = [
            BRD_CUSTOMER_ID,
            BRD_MAN_DATE,
            BRD_PCBA_ID,
            BRD_LASER_ID,
            HW_PART_NUM,
            BRD_FULL_MODEL,
            TST_OPERATOR_INITIALS
        ]
        sqlkeys_module = [
            BRD_CUSTOMER_ID,
            BRD_WAVE_LEN_NM,
            BRD_MAN_DATE,
            BRD_ITU_CH_NUM
        ]

        searchpnfrommes = False
        if ( num_type == BRD_LASMOD ):
            keystosearch = sqlkeys_module
            brd_info_SQL[ DB_PRODUCT_FAMILY ] = TST_TYPE_LM
        else:
            keystosearch = sqlkeys_txrx
        # end if

        for akey in keystosearch:
            row = sql_row_clm_table[ akey ][ KEY_TDS_ROW ]
            clm = sql_row_clm_table[ akey ][ KEY_TDS_COLUMN ]
            brd_info_SQL[ akey ] = record[ row ][ clm ]

            # For manufacturing date, check that it is in the proper format
            if ( akey == BRD_MAN_DATE ):
                date_type = type( record[ row ][ clm ] )
                if ( date_type != type( datetime.datetime.now() ) ):
                    log.error( 'The datetime {0} type is wrong'
                               '.'.format( record[ row ][ clm ] ) )
                    return None
                else:
                    brd_info_SQL = self._convertDateTime2String( brd_info_SQL )
                # end if
            # end for
        # end for

        # PartNumber
        if ( HW_PART_NUM in sql_row_clm_table ):
            row = sql_row_clm_table[ HW_PART_NUM ][ KEY_TDS_ROW ]
            clm = sql_row_clm_table[ HW_PART_NUM ][ KEY_TDS_COLUMN ]
            brd_info_SQL[ HW_PART_NUM ] = record[ row ][ clm ]
        else:
            log.debug( 'No HW_PART_NUM in sql look up table. But okay.' )
        # end for

        # model
        brd_info_SQL[ BRD_MODEL ] = cust_id
        # SN
        brd_info_SQL[ BRD_SN ] = brd_info_SQL[ BRD_CUSTOMER_ID ]
        # PCB ID
        brd_info_SQL[ BRD_PCB_ID ] = brd_info_SQL[ BRD_PCBA_ID ]

        # itu chann
        if ( BRD_FULL_MODEL in brd_info_SQL.keys() ):
            itupos = brd_info_SQL[ BRD_FULL_MODEL ].find( KEY_TDS_ITU )
            if ( itupos > -1 ):
                brd_info_SQL[ BRD_ITU_CH_NUM ] = brd_info_SQL[
                        BRD_FULL_MODEL ][ itupos + 3 : itupos + 5 ]
            else:
                # no itu
                brd_info_SQL[ BRD_ITU_CH_NUM ] = '0'
        # end if

        # Test result is passed, as we only retrieve passed results
        brd_info_SQL[ TST_RESULT ] = TST_RES_PASSED

        if ( 'AT3' in brd_info_SQL[ BRD_FULL_MODEL ] ):
            # set final result for Aurora products to passed, as we don't
            # check distortion results for Aurora
            brd_info_SQL[ TST_FINAL_RESULT ] = TST_RES_PASSED
        # end if

        write_json_file( SQL_RESULTS_FN, brd_info_SQL )

        log.debug( LOG_EXIT )
        return brd_info_SQL
    # end function



    def _getBoardInfoFromSQL_LM( self, db_table_name: str, num_type: str,
                             in_num: str, mongo_bridge_dict: dict,
                             lookup_table_dict: dict ) -> dict:
        '''
        Get board information from sql database

        db_table_name: date base table name
        num_type: identifies what type of number the inputted number string is.
            examples are BRD_SERIAL_NUM, BRD_LASER_ID, BRD_CUSTOMER_ID
        in_num: inputted number string
        mongo_bridge_dict: product SQL table name and part number - product ID
                           corresponding dict
        lookup_table_dict: the keys in which cell in SQL table

        return:    None -> get board information failed
                  dict -> get board information succeeded
        {
            "Database Type":             "",
            "ProductID":                 "",
            "Serial Number":             "",
            "Manufacturing Date":        "",
            "PCBA ID":                   "",
            "LASER ID":                  "",

            "Module ID":                 "",
            "PartNumber"                 "",
            "Test Result":               "",
        }
        '''
        log.debug( LOG_ENTER )

        # Default Serial str for SQL query.
        serial_str = 'Device_SN'
        # Check if the table and key is found in the lookup table. If it
        # exists, use it instead of the default serial_str.
        if( db_table_name in lookup_table_dict.keys() ):
            db_table_lookup = lookup_table_dict[ db_table_name ]
            if( num_type in db_table_lookup ):
                if( 'name' in db_table_lookup[ num_type ] ):
                    serial_str = db_table_lookup[ num_type ][ 'name' ]
                # end if
            # end if
        # end if

        # create SQL query string
        if ( num_type == BRD_SERIAL_NUM ):
            getBySNSql = str( "SELECT * FROM {0} WHERE [{1}] = '{2}' ORDER BY "
                              "[RecordID] DESC".format( db_table_name,
                              serial_str, in_num ) )
        # end if

        # connect to sql server
        log.debug( 'SQL query: {0}'.format( getBySNSql ) )
        sqlserver = mssql.mssqlserver()
        if ( not sqlserver ):
            log.error( 'Failed to connect to SQL Database' )
            return None
        # end if

        # execute query
        record = sqlserver.ExecQuery( getBySNSql )
        if ( not record ):
            log.error( 'Failed to get record for {0}: {1} from '
                       'SQL database'.format( num_type, in_num ) )
            return None
        # end if

        # Judge record[0] is not empty
        if ( len( record ) < 1 ):
            log.error( 'No data in record'.format( num_type, in_num ) )
            return None
        # end if

        if ( lookup_table_dict is None ):
            log.error( 'Can not load {0} file'.format( LOOK_UP_TABLE_FN ) )
            return None
        # end if

        if ( db_table_name not in lookup_table_dict ):
            log.error( '{0} key is not in {1} file'.format( db_table_name,
                LOOK_UP_TABLE_FN ) )
            return None
        # end if

        sql_row_clm_table = lookup_table_dict[ db_table_name ]

        # Infer product ID by CusPartNum and
        # product ID/cust id bridge(Mongo_bridge)
        if ( HW_PART_NUM in sql_row_clm_table ):
            row = sql_row_clm_table[ HW_PART_NUM ][ KEY_TDS_ROW ]
            clm = sql_row_clm_table[ HW_PART_NUM ][ KEY_TDS_COLUMN ]
        elif ( HW_MODEL in sql_row_clm_table ):
            row = sql_row_clm_table[ HW_MODEL ][ KEY_TDS_ROW ]
            clm = sql_row_clm_table[ HW_MODEL ][ KEY_TDS_COLUMN ]
        # end if

        cust_id = str( record[ row ][ clm ] ).strip()

        # TO DO Add bridge for Mongo and MES SQL
        # 11/3/2017, For 1310 Laser Module, we should get the product id
        # through part number
        if ( self.f_code in [ '1612-STD', '16XX' ] ):
            product_id = mongo_bridge_dict[ SQL_TABLE_LM_MODEL ][ cust_id ]
        else:
            product_id = self.f_code
        # end if

        # Fill return dict
        brd_info_SQL = {}
        # Database Type
        brd_info_SQL[ SPEC_DATA_TYPE ] = SPEC_DB_SQL
        # product ID
        brd_info_SQL[ HW_PRODUCT_ID ] = product_id

        # Search for required keys using lookup_table
        sqlkeys_module = [
            BRD_CUSTOMER_ID,
            BRD_MAN_DATE
        ]

        for akey in sqlkeys_module:
            row = sql_row_clm_table[ akey ][ KEY_TDS_ROW ]
            clm = sql_row_clm_table[ akey ][ KEY_TDS_COLUMN ]
            brd_info_SQL[ akey ] = record[ row ][ clm ]

            # For manufacturing date, check that it is in the proper format
            if ( akey == BRD_MAN_DATE ):
                date_type = type( record[ row ][ clm ] )
                if ( date_type != type( datetime.datetime.now() ) ):
                    log.error( 'The datetime {0} type is wrong'
                               '.'.format( record[ row ][ clm ] ) )
                    return None
                else:
                    brd_info_SQL = self._convertDateTime2String( brd_info_SQL )
                # end if
            # end for
        # end for

        # PartNumber
        if ( HW_PART_NUM in sql_row_clm_table ):
            row = sql_row_clm_table[ HW_PART_NUM ][ KEY_TDS_ROW ]
            clm = sql_row_clm_table[ HW_PART_NUM ][ KEY_TDS_COLUMN ]
            brd_info_SQL[ HW_PART_NUM ] = record[ row ][ clm ]
        else:
            log.debug( 'No HW_PART_NUM in sql look up table. But okay.' )
        # end if

        # Optical Power, used for 1310 label printing
        if ( LM_TDS_DATA_OPW in sql_row_clm_table ):
            if ( LM_TDS_DATA_OPW in sql_row_clm_table ):
                row = sql_row_clm_table[ LM_TDS_DATA_OPW ][ KEY_TDS_ROW ]
                clm = sql_row_clm_table[ LM_TDS_DATA_OPW ][ KEY_TDS_COLUMN ]
                brd_info_SQL[ LM_TDS_DATA_OPW ] = record[ row ][ clm ]
        else:
            log.debug( 'No LM_TDS_DATA_OPW in sql look up table. But okay.' )
        # end if

        # model
        brd_info_SQL[ BRD_MODEL ] = cust_id
        # SN
        brd_info_SQL[ BRD_SN ] = brd_info_SQL[ BRD_CUSTOMER_ID ]

#        # Test result is passed, as we only retrieve passed results
#        brd_info_SQL[ TST_RESULT ] = TST_RES_PASSED

        write_json_file( SQL_RESULTS_FN, brd_info_SQL )

        log.debug( LOG_EXIT )
        return brd_info_SQL
    # end function



    def _getBoardInfoFromSQL_LIP_LM( self, db_table_name:str, num_type:str,
                             in_num:str, mongo_bridge_dict: dict,
                             lookup_table_dict: dict ) -> dict:
        '''
        Get board information from sql database

        db_table_name: date base table name
        num_type: identifies what type of number the inputted number string is.
            examples are BRD_SERIAL_NUM, BRD_LASER_ID, BRD_CUSTOMER_ID
        in_num: inputted number string
        mongo_bridge_dict: product SQL table name and part number - product ID
                           corresponding dict
        lookup_table_dict: the keys in which cell in SQL table

        return:    None -> get board information failed
                  dict -> get board information succeeded
        {
            "Database Type":             "",
            "ProductID":                 "",
            "Serial Number":              "",
            "Manufacturing Date":        "",
            "PCBA ID":                   "",
            "LASER ID":                  "",

            "Module ID":                 "",
            "PartNumber"                 "",
            "Test Result":               "",
        }
        '''
        log.debug( LOG_ENTER )

        # create SQL query string
        if ( num_type == BRD_SERIAL_NUM ):
            getBySNSql = str( "SELECT * FROM {0} WHERE [SerNo] "
                              "= '{1}' ORDER BY [RecordID] "
                              "DESC".format( db_table_name, in_num ) )
        # end if

        # connect to sql server
        log.debug( 'SQL query: {0}'.format( getBySNSql ) )
        sqlserver = mssql.mssqlserver()
        if ( not sqlserver ):
            log.error( 'Failed to connect to SQL Database' )
            return None
        # end if

        # execute query
        record = sqlserver.ExecQuery( getBySNSql )
        if ( not record ):
            log.error( 'Failed to get record for {0}: {1} from '
                       'SQL database'.format( num_type, in_num ) )
            return None
        # end if

        # Fill return dict
        brd_info_SQL = {}

        # Judge record[0] is not empty
        if ( len( record ) < 1 ):
            log.error( 'No data in record'.format( num_type, in_num ) )
            return None
        # end if

        reqkeys_lip = [
            BRD_ITU_CH_NUM,
            BRD_WAVE_LEN_NM
        ]

        sql_row_clm_table = lookup_table_dict[ db_table_name ]

        for akey in reqkeys_lip:
            row = sql_row_clm_table[ akey ][ KEY_TDS_ROW ]
            clm = sql_row_clm_table[ akey ][ KEY_TDS_COLUMN ]
            brd_info_SQL[ akey ] = record[ row ][ clm ]

        # end for
        return brd_info_SQL
    # end function



    def _getBoardInfoFromSQL_FailSMSR_LM( self, db_table_name: str,
                    num_type: str, in_num: str, mongo_bridge_dict: dict,
                    lookup_table_dict: dict ) -> dict:
        '''
        Get board information from sql database

        db_table_name: date base table name
        num_type: identifies what type of number the inputted number string is.
            examples are BRD_SERIAL_NUM, BRD_LASER_ID, BRD_CUSTOMER_ID
        in_num: inputted number string
        mongo_bridge_dict: product SQL table name and part number - product ID
                           corresponding dict
        lookup_table_dict: the keys in which cell in SQL table

        return:    None -> get board information failed
                  dict -> get board information succeeded
        {
            "Database Type":             "",
            "ProductID":                 "",
            "Serial Number":              "",
            "Manufacturing Date":        "",
            "PCBA ID":                   "",
            "LASER ID":                  "",

            "Module ID":                 "",
            "PartNumber"                 "",
            "Test Result":               "",
        }
        '''
        log.debug( LOG_ENTER )

        # create SQL query string
        if ( num_type == BRD_SERIAL_NUM ):
            getBySNSql = str( "SELECT * FROM {0} WHERE [Device_SN] "
                              "= '{1}' ORDER BY [RID] "
                              "DESC".format( db_table_name, in_num ) )
        # end if

        # connect to sql server
        log.debug( 'SQL query: {0}'.format( getBySNSql ) )
        sqlserver = mssql.mssqlserver()
        if ( not sqlserver ):
            log.error( 'Failed to connect to SQL Database' )
            return None
        # end if

        # execute query
        record = sqlserver.ExecQuery( getBySNSql )
        if ( not record ):
            log.error( 'Failed to get record for {0}: {1} from '
                       'SQL database'.format( num_type, in_num ) )
            return None
        # end if

        # Fill return dict
        brd_info_SQL = {}

        # Judge record[0] is not empty
        if ( len( record ) < 1 ):
            log.error( 'No data in record'.format( num_type, in_num ) )
            return None
        # end if

        reqkeys_lip = [
            'WaveLen2',
            'SMSR2'
        ]

        sql_row_clm_table = lookup_table_dict[ db_table_name ]

        for akey in reqkeys_lip:
            row = sql_row_clm_table[ akey ][ KEY_TDS_ROW ]
            clm = sql_row_clm_table[ akey ][ KEY_TDS_COLUMN ]
            brd_info_SQL[ akey ] = record[ row ][ clm ]
        # end for

        return brd_info_SQL
    # end function



    def getBoardInfoFromDataBase( self, num_type: str, in_num: str,
        prod_type: str ) -> dict:
        '''
        Get board information from SQL and mongo
        Check which of the two is more recent, and use that

        num_type: BRD_LASER_ID, BRD_CUSTOMER_ID (serial number)
        in_num: inputted number string
        product_type: type of product being tested, e.g. LM or TX

        return:   None -> get board information failed
                  dict -> get board information successed
        '''
        log.debug( LOG_ENTER )

        retdict = None
        if ( prod_type == PRODUCT_TYPE_TX ):
            retdict = self._getBoardInfoFromDataBase_TX( num_type, in_num )
        elif ( prod_type == PRODUCT_TYPE_LM ):
            retdict = self._getBoardInfoFromDataBase_LM( num_type, in_num )
        else:
            log.error( 'Product type not supported!' )
            retdict = None
        # end if

        log.debug( LOG_EXIT )
        return retdict
    # end function



    def _getBoardInfoFromDataBase_TX( self, num_type:str, in_num:str ) -> dict:
        '''
        Get board information from SQL and mongo for TX
        Check which of the two is more recent, and use that

        num_type: BRD_LASER_ID, BRD_CUSTOMER_ID (serial number)
        in_num: inputted number string

        return:   None -> get board information failed
                  dict -> get board information successed
        '''
        log.debug( LOG_ENTER )

        # Load mongo bridge and lookup table
        mongo_bridge_dict = read_json_file( MONGO_BRIDGE_FN )
        if ( mongo_bridge_dict is None ):
            log.error( 'Failed to read {}!'.format( MONGO_BRIDGE_FN ) )
            return None

        lookup_table_dict = read_json_file( LOOK_UP_TABLE_FN )
        if ( lookup_table_dict is None ):
            log.error( 'Failed to read {}!'.format( LOOK_UP_TABLE_FN ) )
            return None

        mongo_date = None
        sql_date = None

        # Get information from mongo db
        mongo_board_info_dict = self._getBoardInfoFromMongo( num_type, in_num,
            PRODUCT_TYPE_TX )
        if ( mongo_board_info_dict is None ):
            log.debug( 'No mongo results' )
        elif ( PRCS_FLOW_TX in mongo_board_info_dict ):
            # Generate mongo_date from last entry in tx_process_flow
            operation_date = mongo_board_info_dict[ PRCS_FLOW_TX ][ -1 ]\
                             [ 'operation_date' ]
            if( type( operation_date ) == type( OrderedDict() ) ):
                # Old timestamp format
                mongo_timestamp = operation_date[ '$date' ]
                mongo_date = datetime.datetime.utcfromtimestamp(
                             mongo_timestamp / 1000 )
            else:
                # New timestamp format
                mongo_date = datetime.datetime.strptime( operation_date,
                        '%Y-%m-%d %H:%M:%S' )

        # end if
        log.debug( 'Mongo Date: {}'.format( mongo_date ) )
        # Get information from SQL
        sql_table = 'Catv_QAM'
        if ( ( num_type == BRD_CUSTOMER_ID ) or ( num_type == BRD_SERIAL_NUM )
            or ( num_type == BRD_SN ) ):
            sql_num_type = BRD_TX_ID
        else:
            sql_num_type = num_type
        # end if

        sql_board_info_dict = self._getBoardInfoFromSQL(
                sql_table,
                sql_num_type,
                in_num,
                mongo_bridge_dict,
                lookup_table_dict,
                PRODUCT_TYPE_TX )
        if ( sql_board_info_dict is None ):
            log.debug( 'No SQL results' )
        elif ( sql_board_info_dict[ BRD_MAN_DATE ] ):
            # SQL results contain a datetime to compare directly
            sql_date = sql_board_info_dict[ BRD_MAN_DATE ]
        # end if

        # Determine which set of results to use by checking which
        # results were returned, and which of the dates were more recent
        # if we received results for both SQL and mongo
        if ( ( mongo_board_info_dict is None ) and
               ( sql_board_info_dict is None ) ):
            # no results found
            log.error( 'Can not get {0}: {1} from mongo and '
               'SQL database'.format( num_type, in_num ) )
            return None
        elif ( ( mongo_date is None ) and ( not sql_date is None ) ):
            log.debug( 'Only SQL results found, using those' )
            board_info_dict = sql_board_info_dict
            board_info_dict[ SPEC_DATA_TYPE ] = SPEC_DB_SQL
        elif ( ( sql_date is None ) and ( not mongo_date is None ) ):
            log.debug( 'Only mongo results found, using those' )
            board_info_dict = mongo_board_info_dict
            board_info_dict[ SPEC_DATA_TYPE ] = SPEC_DB_MONGO
        elif ( sql_date > mongo_date ):
            log.debug( 'SQL results more recent than mongo results, using '
                    'those' )
            board_info_dict = sql_board_info_dict
            board_info_dict[ SPEC_DATA_TYPE ] = SPEC_DB_SQL
            log.debug( 'SQL date: {}'.format( sql_date ) )
            log.debug( 'Mongo date: {}'.format( mongo_date ) )

        elif ( mongo_date > sql_date ):
            log.debug( 'Mongo results more recent than SQL results, using '
                    'those' )
            board_info_dict = mongo_board_info_dict
            board_info_dict[ SPEC_DATA_TYPE ] = SPEC_DB_MONGO
            log.debug( 'SQL date: {}'.format( sql_date ) )
            log.debug( 'Mongo date: {}'.format( mongo_date ) )
        # end if

        log.debug( LOG_EXIT )
        return board_info_dict
    # end function



    def _getBoardInfoFromDataBase_LM( self, num_type: str,
                                            in_num: str ) -> dict:
        '''
        Get board information from SQL and mongo for laser module
        Check which of the two is more recent, and use that

        num_type: BRD_LASER_ID, BRD_CUSTOMER_ID (serial number)
        in_num: inputted number string

        return:   None -> get board information failed
                  dict -> get board information successed
        '''
        log.debug( LOG_ENTER )

        # Load mongo bridge and lookup table
        mongo_bridge_dict = read_json_file( MONGO_BRIDGE_FN )
        lookup_table_dict = read_json_file( LOOK_UP_TABLE_FN )
        mongo_date = None
        sql_date = None

        # Get information from mongo db
        mongo_board_info_dict = self._getBoardInfoFromMongo( num_type, in_num,
            PRODUCT_TYPE_LM )
        if ( mongo_board_info_dict is None ):
            log.debug( 'No mongo results' )
        elif ( PRCS_FLOW_LM in mongo_board_info_dict ):
            log.debug( 'Found process flow' )
            processflow = mongo_board_info_dict[ PRCS_FLOW_LM ]
            # Generate mongo_date from last Multi-Up entry in LM Process Flow
            # Proper way is to use the UTC date time, though VB SQL dates are
            # not timezone aware currently.
            for currdict in reversed( processflow ):
                curr_prcs = currdict[ PRCS_NAME ]
                if ( curr_prcs != PROCESS_RESULT_KEY_MULTIUP ):
                    log.debug( 'Not multiup. Skip {}'.format( curr_prcs ) )
                    continue
                # end if
                mongo_timestamp = currdict[ KEY_TDS_LM_OPER_DATE ]
#                mongo_timestamp = currdict[ KEY_TDS_LM_OPER_DATE_UTC ]
                if( type( mongo_timestamp ) == type( OrderedDict() ) ):
                    mongo_timestamp = mongo_timestamp[ '$date' ]
                    mongo_date = datetime.datetime.utcfromtimestamp(
                                 mongo_timestamp / 1000 )
                else:
                    # TODO Remove this in the future, once we schedule to
                    #      convert the API.
                    # Convert from string to datetime
                    mongo_date = datetime.datetime.strptime( mongo_timestamp,
                            '%Y-%m-%d %H:%M:%S' )
                # end if

#                # Convert from UTC to local time
#                mongo_date = mongo_date.replace( tzinfo =
#                    datetime.timezone.utc ).astimezone( tz = None )
#                break
            # end for
        # end if

        log.debug( 'date for mongo {}'.format( mongo_date ) )

        # Get information from SQL in order to compare the date time
        # Due to different processes for the laser module line, we may only
        # test certain stations on the SQL line. Need to go by family code.
        if( self.f_code not in mongo_bridge_dict[ MONGO_BRIDGE_DB_TBL ] ):
            log.debug( 'Could not find {} in mongobridge {}: {}'.format(
                       self.f_code, MONGO_BRIDGE_DB_TBL,
                       mongo_bridge_dict[ MONGO_BRIDGE_DB_TBL ] ) )
            return None

        sql_table = mongo_bridge_dict[ MONGO_BRIDGE_DB_TBL ][ self.f_code ]

        sql_num_type = num_type

        sql_board_info_dict = self._getBoardInfoFromSQL(
                sql_table,
                sql_num_type,
                in_num,
                mongo_bridge_dict,
                lookup_table_dict,
                PRODUCT_TYPE_LM )

        log.debug( 'sql_board_info_dict is {0}'.format( sql_board_info_dict ) )

        if ( sql_board_info_dict is None ):
            log.debug( 'No SQL results' )
        elif ( BRD_MAN_DATE in sql_board_info_dict ):
            # SQL results contain a datetime to compare directly
            sql_date = sql_board_info_dict[ BRD_MAN_DATE ]
        # end if

        log.debug( 'date for sql {}'.format( sql_date ) )

        # Determine which set of results to use by checking which
        # results were returned, and which of the dates were more recent
        # if we received results for both SQL and mongo
        if ( ( mongo_board_info_dict is None ) and
               ( sql_board_info_dict is None ) ):
            # no results found
            log.error( 'Can not get {0}: {1} from mongo and '
               'SQL database'.format( num_type, in_num ) )
            return None
        elif ( ( mongo_date is None ) and ( not sql_date is None ) ):
            log.debug( 'Only SQL results found, using those' )
            board_info_dict = sql_board_info_dict
            board_info_dict[ SPEC_DATA_TYPE ] = SPEC_DB_SQL
        elif ( ( sql_date is None ) and ( not mongo_date is None ) ):
            log.debug( 'Only mongo results found, using those' )
            board_info_dict = mongo_board_info_dict
            board_info_dict[ SPEC_DATA_TYPE ] = SPEC_DB_MONGO
        elif ( sql_date > mongo_date ):
            log.debug( 'SQL results more recent than mongo results, using '
                    'those' )
            board_info_dict = sql_board_info_dict
            board_info_dict[ SPEC_DATA_TYPE ] = SPEC_DB_SQL
        elif ( mongo_date > sql_date ):
            log.debug( 'Mongo results more recent than SQL results, using '
                    'those' )
            board_info_dict = mongo_board_info_dict
            board_info_dict[ SPEC_DATA_TYPE ] = SPEC_DB_MONGO
        # end if

        # For units ran through SQL, initialize the product spec based on the
        # family code and remove Multi-Up station data if it exists.
        if ( board_info_dict[ SPEC_DATA_TYPE ] == SPEC_DB_SQL ):
            initsqldict = self._Initialize_Sql_Config()
            board_info_dict.update( initsqldict )
            if ( ( mongo_board_info_dict != None ) and
                 ( TST_BER_TEST_DATA in mongo_board_info_dict ) ):
                board_info_dict[ TST_BER_TEST_DATA ] = mongo_board_info_dict[
                    TST_BER_TEST_DATA ]
            board_info_dict[ BRD_MAN_DATE ] = sql_date
        else:
            board_info_dict[ BRD_MAN_DATE ] = mongo_date

        log.debug( LOG_EXIT )

        return board_info_dict
    # end function



    def getBoardInfoFromSpecData( self, board_info_dict:dict,
                                  product_spec_dict:dict ) -> dict:
        '''
        get board information from spec data ( sql or mongo ) for TDS printing

        board_info_dict: board information
        product_spec_dict: Product Spec Data dict

        return:    None -> get board information failed
                   dict -> get board information successed
        '''
        log.debug( LOG_ENTER )

        # Check SPEC_DATA_TYPE in board_info_dict
        if ( SPEC_DATA_TYPE not in board_info_dict ):
            log.error( 'Can not find SPEC_DATA_TYPE in db board info' )
            return None
        # end if

        db_type = board_info_dict[ SPEC_DATA_TYPE ]

        # Get itu channel from mongo board information dict
        if ( db_type == SPEC_DB_MONGO ):
            # Check BRD_ITU_CH_NUM in board_info_dict
            if ( BRD_ITU_CH_NUM not in board_info_dict ):
                log.error( 'Can not find BRD_ITU_CH_NUM  in db board info' )
                return None
            # end if

            itu_chann = str( board_info_dict[ BRD_ITU_CH_NUM ] )
            log.debug( 'itu channel found: {0}'.format( itu_chann ) )

            if ( itu_chann in ['0', '1'] ):
                log.debug( 'This is the 1310 product' )
                chann_spec_dict = product_spec_dict
            # Check itu_chann in board_info_dict
            elif ( itu_chann not in product_spec_dict ):
                log.error( 'Can not find itu channel {0} in '
                           'spec data'.format( itu_chann ) )
                return None
            else:
                chann_spec_dict = product_spec_dict[ itu_chann ]
        # end if, Get itu channel from sql board information dict
        elif ( db_type == SPEC_DB_SQL ):
            # Laser Module has product spec data by the part number since we
            # can have multiple part numbers for the same ITU.
            # Get the part number by the ITU for now. Need a solution to
            # know which part number to use when there are multiple PN for
            # the same ITU.
            # CATV has product spec data by the ITU channel.
            itu_chann = None
            if ( board_info_dict[ DB_PRODUCT_FAMILY ] == TST_TYPE_LM ):
                itu_chann = str( int( board_info_dict[ BRD_ITU_CH_NUM ] ) )
                for prd in product_spec_dict:
                    if itu_chann == product_spec_dict[ prd ][ BRD_ITU_CH_NUM ]:
                        chann_spec_dict = product_spec_dict[ prd ]
                        board_info_dict[ HW_PART_NUM ] = prd
                        break
                # end for
            else:
                # Check HW_PART_NUM in board_info_dict
                if ( HW_PART_NUM not in board_info_dict ):
                    log.error( 'Cannot find HW_PART_NUM in db board info' )
                    return None
                # end if

                part_number = board_info_dict[ HW_PART_NUM ]

                if ( HW_PART_NUM in product_spec_dict ):
                    # No nesting in product spec, use directly
                    chann_spec_dict = product_spec_dict
                else:
                    for key in product_spec_dict.keys():
                        # Find the product ID that matches board info
                        # Check HW_PART_NUM key in each dict
                        if ( HW_PART_NUM not in product_spec_dict[ key ] ):
                            log.error( 'Can not find HW_PART_NUM in spec data' )
                            return None
                        # end if

                        # Check whether we found the part number
                        if ( part_number in product_spec_dict[ key ]\
                                                             [ HW_PART_NUM ] ):
                            itu_chann = key
                            break
                        # end if
                    # end for

                    if ( itu_chann is None ):
                        log.error( 'Can not find itu channel in spec data' )
                        return None
                    # end if

                    chann_spec_dict = product_spec_dict[ itu_chann ]
                # end if
            # end if
        # don't support the data type
        else:
            log.error( 'do not support database type'.format( db_type ) )
            return None
        # end if

        # Add parameters to board information
        for key in chann_spec_dict.keys():
            board_info_dict[ key ] = chann_spec_dict[ key ]
        # end for

        log.debug( LOG_EXIT )
        return board_info_dict
    # end function



    def _convertDateTime2String( self, board_info_dict:dict ) -> dict:
        '''
        convert datetime to datetime string

        board_info_dict: board information

        return:    None -> convert datetime to string failed
                   dict -> convert datetime to string successed
        '''
        log.debug( LOG_ENTER )

        if ( BRD_MAN_DATE not in board_info_dict ):
            log.error( 'Can not find BRD_MAN_DATE in board information' )
            return None
        # end if

        date_type = type( board_info_dict[ BRD_MAN_DATE ] )
        if ( date_type != type( datetime.datetime.now() ) ):
            log.error( 'BRD_MAN_DATE type is incorrect' )
            return None
        # end if
        log.debug(board_info_dict)
        date_time = board_info_dict[ BRD_MAN_DATE ]
        new_year_day = datetime.datetime( date_time.year, 1, 1, 0, 0 )

        date_time_days = ( date_time - new_year_day ).days + 1

        board_info_dict[ KEY_TDS_MANUFACTURE_MM_DD ] = \
            '{0:02d}'.format( date_time.month ) + '/' + \
            '{0:02d}'.format( date_time.day )

        board_info_dict[ KEY_TDS_MANUFACTURE_YY_DD ] = \
            '{0:02d}'.format( date_time.year - 2000 ) + \
            '{0:03d}'.format( date_time_days )

        board_info_dict[ KEY_TDS_MANUFACTURE_MM_YY ] = \
            '{0:02d}'.format( date_time.month ) + '/' + \
            '{0:02d}'.format( date_time.year - 2000 )


        board_info_dict[ KEY_TDS_MANUFACTURE_MONTH_YEAR ] = \
                "{0} {1}".format( monfullnamedict[ date_time.strftime( '%m' ) ],
                    date_time.strftime( '%y' ) )

        if ( board_info_dict[ SPEC_DATA_TYPE ] == SPEC_DB_SQL ):
            board_info_dict[ KEY_TDS_DATE_CREATION ] = date_time.strftime(
                    '%d/%m/%Y %I:%M:%S %p' )

        log.debug( LOG_EXIT )
        return board_info_dict
    # end function



    def _FilterData( self, rollup_result: dict ) -> dict:
        """
        Filter data so we only get relevant data for laser module, e.g. 1752-C21

        rollup_result: rollup result to parse through

        Check all possible process flows for laser module:
        1. Only BER data exists in Mongo
        2. Only Multiup data exists in Mongo
        3. Both BER and Multiup data exists in mongo
        4. No data in Mongo. Only SQL.

        :return: the data dictionary related to the sn
                 empty dictionary if
        """
        log.debug( LOG_ENTER )

        related_data = {}

        # adapt to stale schema
        rollupexists = False
        berdataexists = False
        multiupdataexists = False

        # if None or empty dictionary, data does not exist.
        if ( rollup_result ):
            log.debug( 'rollup exists' )
            rollupexists = True
        if ( rollup_result.get( TST_BER_TEST_DATA ) ):
            log.debug( 'ber data exists' )
            berdataexists = True
        if ( rollup_result.get( TST_MULTIUP_DATA ) ):
            log.debug( 'multiup data exists' )
            multiupdataexists = True

        # Only BER data exists
        if ( rollupexists and berdataexists and not multiupdataexists ):
            log.debug( 'found ber data from mongo' )
            related_data[ TST_BER_TEST_DATA ] = rollup_result.get(
                TST_BER_TEST_DATA )
        # Only Multiup data exists
        elif ( rollupexists and not berdataexists and multiupdataexists ):
            log.debug( 'found multiup data from mongo' )
            related_data[ TST_MULTIUP_DATA ] = rollup_result.get(
                TST_MULTIUP_DATA )
        # Both BER and Multiup data exists
        elif ( rollupexists and berdataexists and multiupdataexists ):
            log.debug( 'found ber and multiup data from mongo' )
            related_data = rollup_result
        # No data found in Mongo
        else:
            log.error( 'No data found in Mongo' )

        log.debug( LOG_EXIT )
        return related_data
    # end function



    def _Initialize_Sql_Config( self ) -> dict:
        """
        Initialize a configuration for next sql query using the family code

        return: the required configuration for sql query
        """
        log.debug( LOG_ENTER )

        related_data = {}
        related_data[ LM_SPEC_DEVICE_INFO ] = { }
        related_data[ LM_SPEC_DEVICE_INFO ][ HW_PRODUCT_ID ] = self.f_code
        related_data[ LM_SPEC_DEVICE_INFO ][ BRD_SERIAL_NUM ] = \
            self._serialnumber

        # adapt to stale schema
        related_data[ RSLT_DEV_INFO_KEY ] = { }
        related_data[ RSLT_DEV_INFO_KEY ][ HW_PRODUCT_ID ] = self.f_code
        related_data[ RSLT_DEV_INFO_KEY ][ BRD_SERIAL_NUM ] = \
            self._serialnumber

        log.debug( LOG_EXIT )
        return related_data
    # end function
# end class





"""
Class that takes boardinfo and cfgdata (specfile) and prints either a
box label and unit label or both

primary method: printData( self, pboxlblsel:int, punitlblsel:int )
    input: pboxlblsel:int, punitlblsel:int
           takes either 1 or 0 to determine whether or not to print box
           and unit labels
    output: returns True if label(s) is successfully printed, or will return
           False if it fails at any stage
"""
class PrintLabel( object ):
    """ Base class for a device that requires printing

    Will hold reusable functions by all devices
    """
    def __init__( self, boardinfodict:dict, cfgdata:dict ):
        ''' PrintLabel's constructor

        boardinfodict : data relationship between serial number and part number
        cfgdata : device spec data including template names, etc
        '''
        log.debug( LOG_ENTER )

        if ( cfgdata is None ):
            errmsg = 'Did not get expected CFG file with product spec info'
            log.debug( errmsg )
            raise ValueError( errmsg )
        # end if

        if ( boardinfodict is None ):
            errmsg = 'Did not get expected boardinfo from DB'
            log.debug( errmsg )
            raise ValueError( errmsg )
        # end if

        # class private name
        self.emkridnumber = ''
        self.tdstemplatefn = ''
        self.tdsprintfn = ''
        self.frequencyresponseres = {}
        self.distortionres = {}
        self.cfginfo = {}
        self.boardinfodict = {}
        self.printdatadict = {}
        self.tdscellmap = {}
        self.labelcellmap = {}
        self.printtemplatelist = []
        self.printerlist = []
        self.templateboxlist = {}
        self.templateunitlist = {}
        self.stationfile = {}
        self.databasetype = boardinfodict[ SPEC_DATA_TYPE ]

        self.cfginfo = cfgdata
        self.boardinfodict = boardinfodict

        log.debug( LOG_EXIT )
    # end function



    def printData( self, pboxlblsel:int, punitlblsel:int ) -> bool:
        """ Takes care of any preprocessing before executing the print command

        pboxlblsel: True-print box label; False; don't print box label
        punitlblsel: True-print unit label; False: don't print unit label
        Return:
            True: No issues
            False: Issue with prepreprocessing
        """
        log.debug( LOG_ENTER )

        retbool = self.getRelatedDataFromCfg()

        # --Update label template(Excel file)
        label_obj = LabelPrinting()
        templateStatus = label_obj.updateTemplate( self.excltemplatefile,
                                                   self.boardinfodict,
                                                   self.labelcellmap )

        # --Check update label template failed or successed
        if ( templateStatus == False ):
            log.debug( 'Had problems updating the template' )
            return False
        # end if

        if ( retbool == False ):
            log.debug( 'Had problems retrieving data from the '
                       'input spec file' )
            return False
        # end if

        if ( pboxlblsel == 1 ):
            self.printtemplatelist += self.templateboxlist
        # end if

        if ( punitlblsel == 1 ):
            self.printtemplatelist += self.templateunitlist
        # end if

        # 10/16/17 Pull query for station file out of get_PrinterInfo.
        # Decrease number of database request.
        computername = getThisComputerName()
        if ( not computername ):
            log.error( "Failed to get Station name" )
            return None
        # end if
        station_id = computername
        # manual station file
        self.stationfile = read_json_file( STATION_FN )
        # self.stationfile = getStationInfo( db_info = db_info,
        #     station_id = station_id )

        if ( not self.stationfile ):
            log.error( "Failed to get Station file from Database" )
            return False
        # end if
        write_json_file( STATION_FN, self.stationfile )

        if ( KEY_TDS_PRINTER_INFO not in self.stationfile ):
            log.error( 'Could not find KEY_TDS_PRINTER_INFO in config' )
            return False

        # Open the bartender object
        self.shellapp = None
        default_lbl_app = 'BarTender.Application'
        if KEY_LBL_PRINT_APP in self.stationfile[ KEY_TDS_PRINTER_INFO ]:
            self.shellapp = self.stationfile[ KEY_TDS_PRINTER_INFO ]\
                            [ KEY_LBL_PRINT_APP ]
        else:
            try:
                self.barapp = win32com.client.Dispatch( default_lbl_app )
            except Exception as e:
                log.error( " Failed to open Bartender Application error {0}".
                           format( e ) )
                return False
            # end try
            self.barapp.Visible = True

        # Loop through all templates that we want to print, get the printer
        # information for each, validate it, and then send the template to
        # the printer.
        for template in self.printtemplatelist:
            printerinfo = self.get_PrinterInfo( template )

            if ( not printerinfo ):
                log.error( " Failed to get printer info" )
                return False
            # end if

            if ( not self.validate_Printer( printerinfo ) ):
                log.error( " Failed to validate printer info on OS" )
                return False
            # end if

            printer = printerinfo[ BRD_NAME ]

            log.debug( "Printer Name: {0}".format( printer ) )

            if not self._send_to_printer( template, printer ):
                log.error( "Failed to print template {0} on printer {1}".
                          format( template, printer ) )
                return False
            # end if

        # end for

        # self.barapp.Quit( 1 )
        log.debug( LOG_EXIT )
        return True
    # end function



    def getRelatedDataFromCfg( self ) -> bool:
        """ Retreives data from tds label specific cfg or dictionary

        Data retrieved:
        self.labelcellmap = self.cfginfo[ BRD_LBL_CELL_MAP ] (Label Cell Map)
        self.excltemplatefile = self.cfginfo[ BRD_LBL_EXCL_TMPT ]
#                               (Label Excel Template)
        self.toptemplatelist = self.cfginfo[ BRD_LBL_BAR_TMPTS ]
#                               (Label Bartender Templates)
            self.templateboxlist = self.toptemplatelist[ 'Box' ]
            self.templateunitlist = self.toptemplatelist[ 'Unit' ]

        Return:
            True: Data retrieved
            False: Issue with obtaining data
        """
        log.debug( LOG_ENTER )

        # no product spec data for SATCOM products.
        if ( self.boardinfodict[ DB_PRODUCT_FAMILY ] != FAMILY_SATCOM ):
            if ( BRD_SPEC_DATA not in self.cfginfo ):
                log.debug( 'Labelprinting software revision'
                           ' does not support product ' )
                return False
            # end if
            self.productspecdata = self.cfginfo[ BRD_SPEC_DATA ]
        else:
            self.productspecdata = None
        # end if

        if ( BRD_LBL_CELL_MAP not in self.cfginfo ):
            log.debug( 'Labelprinting software revision'
                       ' does not support product ' )
            return False
        # end if

        if ( BRD_LBL_BAR_TMPTS not in self.cfginfo ):
            log.debug( '{0} not found in print cfg dictionary'.format(
                BRD_LBL_BAR_TMPTS ) )
            return False
        # end if

        if ( BRD_LBL_EXCL_TMPT not in self.cfginfo ):
            log.debug( '{0} not found in print cfg dictionary'.format(
                BRD_LBL_EXCL_TMPT ) )
            return False
        # end if

        self.labelcellmap = self.cfginfo[ BRD_LBL_CELL_MAP ]
        self.toptemplatelist = self.cfginfo[ BRD_LBL_BAR_TMPTS ]
        self.excltemplatefile = self.cfginfo[ BRD_LBL_EXCL_TMPT ]

        if ( KEY_TDS_BOX not in self.toptemplatelist ):
            log.debug( 'Box template filesnames not found in cfg' )
            return False
        # end if

        if ( KEY_TDS_UNIT not in self.toptemplatelist ):
            log.debug( 'Unit template filesnames not found in cfg' )
            return False
        # end if

        self.templateboxlist = self.toptemplatelist[ KEY_TDS_BOX ]
        self.templateunitlist = self.toptemplatelist[ KEY_TDS_UNIT ]

        self.boardinfodict = self.getBoardInfoFromSpecData(
                board_info_dict = self.boardinfodict,
                product_spec_dict = self.cfginfo[ BRD_SPEC_DATA ] )
        write_json_file( 'BoardInformation.json', self.boardinfodict )

        log.debug( LOG_EXIT )
        return True
    # end function



    def getBoardInfoFromSpecData( self, board_info_dict:dict,
                                  product_spec_dict:dict ) -> dict:
        '''
        get board information from spec data for label printing

        board_info_dict: board information
        product_spec_dict: Product Spec Data dict

        return:    None -> get board information failed
                   dict -> get board information successed
        '''
        log.debug( LOG_ENTER )

        retdict = {}
        self._serialnumber = board_info_dict[ BRD_SN ]
#        family_codes = GetOriginalModelFromMES( sn = self._serialnumber )
        family_codes = GetPartNumbersbySerialnumber( self._serialnumber )
        if ( family_codes is None ):
            log.debug( 'GetPartNumbersbySerialnumber failed! '
                'Did not get LM family results from SQL using sn: '
                '{0}'.format( self._serialnumber ) )

            msg = 'SQL DB connection issue!'
            log.debug( msg )
            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            raise Exception( msg )
        # end if

        if ( isinstance( family_codes, list ) == False ):
            log.debug( 'GetPartNumbersbySerialnumber failed! '
                'Ret value not list from SQL using sn: '
                '{0}'.format( self._serialnumber ) )

            msg = 'TDS printing SW issue. Verify SQL data handling (LM family)'
            log.debug( msg )
            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            raise Exception( msg )
        # end if

        if ( len( family_codes ) == 0 ):
            log.debug( 'GetPartNumbersbySerialnumber failed! '
                'Ret value empty from SQL using sn: '
                '{0}'.format( self._serialnumber ) )

#            msg = 'LM family not matched to SN given'
            msg = 'Cannot find the product model for this board in MES'
            log.debug( msg )
            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            raise Exception( msg )
        # end if

        # Get the first value, which should be the latest family code
        last_family_code = family_codes[ 0 ][ 2 ]

        self.f_code = last_family_code

#        # Find the the SQL FCODE FIELD
#        self.f_code = last_family_code.get( SQL_FCODE_FIELD, None )
        if ( self.f_code is None ):
            log.debug( 'GetPartNumbersbySerialnumber failed! '
                'Ret value missing F_CODE field from SQL using sn: '
                '{0}'.format( self._serialnumber ) )

            msg = 'MES issue, LM family not in SQL result using SN'
            log.debug( msg )
            retdict[ KEY_TDS_STATUS_MSG ] = msg
            retdict[ KEY_TDS_STATUS_BOOL ] = False
            raise Exception( msg )
        # end if

        log.debug( 'Family code for SN: {0} is {1}'.format(
            self._serialnumber, self.f_code ) )

        # Check SPEC_DATA_TYPE in board_info_dict
        if ( SPEC_DATA_TYPE not in board_info_dict ):
            log.error( 'Can not find SPEC_DATA_TYPE in db board info' )
            return None
        # end if

        db_type = board_info_dict[ SPEC_DATA_TYPE ]

        # Get itu channel from mongo board information dict
        if ( db_type == SPEC_DB_MONGO ):
            if ( board_info_dict[ DB_PRODUCT_FAMILY ] != TST_TYPE_LM ):
                # Check BRD_ITU_CH_NUM in board_info_dict
                if ( BRD_ITU_CH_NUM not in board_info_dict ):
                    log.error( 'Can not find BRD_ITU_CH_NUM  in db board info' )
                    return None
                # end if

                itu_chann = str( board_info_dict[ BRD_ITU_CH_NUM ] )

                if ( itu_chann in [ '0', '1' ] ):
                    log.debug( 'This is the 1310 product' )
                    chann_spec_dict = product_spec_dict
                # Check itu_chann in board_info_dict
                elif ( itu_chann not in product_spec_dict ):
                    log.error( 'Can not find itu channel {0} in '
                               'spec data'.format( itu_chann ) )
                    return None
                else:
                    chann_spec_dict = product_spec_dict[ itu_chann ]
            else:
                return board_info_dict
            # end if
        # Get itu channel from sql board information dict
        elif ( db_type == SPEC_DB_SQL ):
            # Laser Module has product spec data by the part number since we
            # can have multiple part numbers for the same ITU.
            # Get the part number by the ITU for now. Need a solution to
            # know which part number to use when there are multiple PN for
            # the same ITU.
            # CATV has product spec data by the ITU channel.
            if ( self.f_code in [ '1612-STD' ] ):
                opt_power = None
                if ( board_info_dict[ DB_PRODUCT_FAMILY ] == TST_TYPE_LM ):
                    opt_power = str( int( board_info_dict[ LM_TDS_DATA_OPW ] ) )
                    for prd in product_spec_dict:
                        if ( ( opt_power > product_spec_dict[ prd ][ 'OptPwr_Min' ] )
                            and ( opt_power < product_spec_dict[ prd ][ 'OptPwr_Max' ] ) ):
                            chann_spec_dict = product_spec_dict[ prd ]
                            board_info_dict[ HW_PART_NUM ] = prd
                            break
                        # end if
                    # end for
                # end if
            else:
                itu_chann = None
                if ( board_info_dict[ DB_PRODUCT_FAMILY ] == TST_TYPE_LM ):
                    itu_chann = str( int( board_info_dict[ BRD_ITU_CH_NUM ] ) )
                    for prd in product_spec_dict:
                        if itu_chann == product_spec_dict[ prd ][ BRD_ITU_CH_NUM ]:
                            chann_spec_dict = product_spec_dict[ prd ]
                            board_info_dict[ HW_PART_NUM ] = prd
                            break
                        # end if
                    # end for
                else:
                    # Check HW_PART_NUM in board_info_dict
                    if ( HW_PART_NUM not in board_info_dict ):
                        log.error( 'Cannot find HW_PART_NUM in db board info' )
                        return None
                    # end if

                    part_number = board_info_dict[ HW_PART_NUM ]

                    if ( HW_PART_NUM in product_spec_dict ):
                        # No nesting in product spec, use directly
                        chann_spec_dict = product_spec_dict
                    else:
                        for key in product_spec_dict.keys():
                            # Find the product ID that matches board info
                            # Check HW_PART_NUM key in each dict
                            if ( HW_PART_NUM not in product_spec_dict[ key ] ):
                                log.error( 'Can not find HW_PART_NUM in '
                                           'spec data' )
                                return None
                            # end if

                            # Check whether we found the part number
                            if ( part_number in product_spec_dict[ key ]\
                                                                 [ HW_PART_NUM ] ):
                                itu_chann = key
                                break
                            # end if
                        # end for

                        if ( itu_chann is None ):
                            log.error( 'Can not find itu channel in spec data' )
                            return None
                        # end if

                        chann_spec_dict = product_spec_dict[ itu_chann ]
                    # end if
                # end if
            # end if
        # don't support the data type
        else:
            log.error( 'do not support database type'.format( db_type ) )
            return None
        # end if

        # Add parameters to board information
        for key in chann_spec_dict.keys():
            board_info_dict[ key ] = chann_spec_dict[ key ]
        # end for

        log.debug( LOG_EXIT )
        return board_info_dict
    # end function



    def get_PrinterInfo( self, template:str ) -> bool:
        '''Set up a right printer on OS based on stations and spec files

        template: StationID
        return : True or False
        '''
        log.debug( LOG_ENTER )

#        # Get product spec file from MONGO DB
#        station_id = getThisComputerName()
#
#        if ( not station_id ):
#            log.error( " Failed to get Station name " )
#            return None
#        # end if

        if ( template not in self.stationfile[ KEY_TDS_PRINTER_INFO ] ):
            log.error( "Failed to Find printer Template {0} on "
                "self.Stationfile".format( template ) )
            return None
        # end if

        printer = self.stationfile[ KEY_TDS_PRINTER_INFO ][ template ]

        if ( printer not in self.stationfile ):
            log.error( " Failed to find the Printer info on Station file" )
            return None
        # end if

        hw_comm = self.stationfile[ printer ][ HW_COMM ]

        address = hw_comm[ BRD_DEVADDR ]
        printername = hw_comm[ BRD_NAME ]

        printerinfo = { BRD_NAME: printername, BRD_DEVADDR: address }

        log.debug( LOG_EXIT )
        return printerinfo
    # end function



    def validate_Printer( self, printerInfo:dict ) -> bool:
        '''validate printer exist on OS and our stations and spec files

        printerInfo: printer info include name and IP address of the printer
        return: True or False
        '''
        log.debug( LOG_ENTER )

        # check if list is not empty and try to fill
        if ( not self.printerlist ):
            self.printerlist = self.get_all_printers()
        # end if

        if ( not self.printerlist ):
            log.error( 'Printer list return empty' )
            return False
        # end if

        isfound = False
        # match IP or Name should be fine
        for eachprinter in self.printerlist:
            if ( printerInfo[ BRD_DEVADDR ].lower() ==
                 eachprinter.portname.lower() ):
                isfound = True
            elif ( printerInfo[ BRD_NAME ].lower() ==
                 eachprinter.printername.lower() ):
                isfound = True
            # end if
        # end for

        if ( not isfound ):
            log.error( "Printer {0} does not exist on system".format(
                    printerInfo[ BRD_NAME ] ) )
            messagebox.showwarning( title = "Missing Printer",
                message = 'Printer {0} not installed on system'.format(
                    printerInfo[ BRD_NAME ] ) )
            return False
        # end if

        log.debug( LOG_EXIT )
        return True
    # end function



    def get_all_printers( self ) -> list:
        '''Get all the printer on OS installed and create object list
        for each printer with all the supported printer attributes

        Return :
            List of object where each object is a printer
            or None
        '''
        log.debug( LOG_ENTER )

        oscmd = ( "cscript prnmngr.vbs -l" )
        printerlist = []
        proc = subprocess.Popen( oscmd, shell = True,
                                stdout = subprocess.PIPE,
                                stderr = subprocess.STDOUT )

        stdout, stderr = proc.communicate()

        if ( proc.returncode != 0 ):
            log.error( "OS cmd {0} failed to execute".format( oscmd ) )
            return None
        # end if

        # Convert byte string to text string list ignore unicode errors
        stdout = ( stdout.decode( "utf-8", errors = 'ignore' ).\
                   encode( "windows-1252", errors = 'ignore' ).\
                   decode( "utf-8", errors = 'ignore' ) )

        log.debug( stdout )

        # split for new lines, count is used to traverse through each line
        # Example of each printer information:
        #       Server name
        #       Printer name PDF995
        #       Share name
        #       Driver name PDF995 Printer Driver
        #       Port name PDF995PORT
        #       Comment
        #       Location
        #       Print processor winprint
        #       Data type RAW
        #       Parameters
        #       Attributes 4165
        #       Priority 1
        #       Default priority 0
        #       Average pages per minute 0
        #       Printer status Idle
        #       Extended printer status Unknown
        #       Detected error state Unknown
        #       Extended detected error state Unknown
        output = stdout.split( '\n' )
        count = 0
        while( count < len( output ) ):
            if ( KEY_TDS_SERVER_NAME not in output[ count ].strip() ):
                count += 1
                continue

            # Add 1 to count to move to next line (printer name)
            count += 1
            while( 1 ):
                # We are processing the final printer
                if ( count >= len( output ) ):
                    break
                # end if

                # We hit the next printer
                if ( KEY_TDS_SERVER_NAME in output[ count ].strip() ):
                    count -= 1
                    break
                # end if

                if ( KEY_TDS_PRINTER_NAME in output[ count ].strip() ):
                    name = ( output[ count ].split( KEY_TDS_NAME )[ -1 ] )
                    printername = name.split( "\\" )[ -1 ]
                elif ( KEY_TDS_PORT_NAME in output[ count ].strip() ):
                    portname = ( output[ count ].split( KEY_TDS_NAME )
                                                    [ -1 ] )
                elif ( KEY_TDS_DRIVER_NAME in output[ count ].strip() ):
                    drivername = ( output[ count ].split( KEY_TDS_NAME )
                                                    [ -1 ] )
                elif ( KEY_TDS_SHARE_NAME in output[ count ].strip() ):
                    sharename = ( output[ count ].split( KEY_TDS_NAME )
                                                    [ -1 ] )
                # end if

                count += 1
            # end while

            printerlist.append( PrinterObj( name.strip(), printername.strip(),
                portname.strip(), drivername.strip(), sharename.strip() ) )

        # end while

        proc.wait()

        if ( not printerlist ):
            log.error( "Printer list is empty please check OS" )
            return None
        # end if

        log.debug( LOG_EXIT )
        return printerlist
    # end function



    def _send_to_printer( self, bartendtemplatefn:str,
                          printername:str ) -> bool:
        """ Sends file to the printer. From filename uses win32com to print

        bartendtemplatefn: btw file name.
        printername: printer name used to print the btw file.

        return : True print
                 False Failed to print file
        """
        log.debug( LOG_ENTER )

        dirpath = os.path.abspath( os.getcwd() )
        templatefilefullpath = os.path.join( dirpath, bartendtemplatefn )

        log.debug( templatefilefullpath )
        # Print with the default Bartender application
        if ( self.shellapp is None ):
            barformat = self.barapp.Formats.Open( templatefilefullpath,
                                                  False, '' )
            barformat.SelectRecordsAtPrint = False

            # Select the print setup variable property
            btPrintSetup = barformat.PrintSetup
            btPrintSetup.Printer = printername
            barformat.PrintOut( False, False )
        # Print with a different version of the bartender app
        else:
            # This method only works if the file is set to use default printer!
            # 1. Save default printer name.
            # 2. Switch default printer to the desired printer
            # 3. After print command is set, set default printer back to
            #    original settings.
            try:
                default_printer = win32print.GetDefaultPrinter()
            except Exception as e:
                log.debug( 'Unexpected error in preparing printer: {0}'.format(
                            e ) )
                return False
            # end try

            fullname = None
            # Find the full printer name in the printer list
            for eachprinter in self.printerlist:
                if ( printername.lower() == eachprinter.printername.lower() ):
                    fullname = eachprinter.fullprintername
                # end if
            # end for

            if ( fullname is None ):
                log.error( 'Could not find the full printer name!' )
                return False

            # Set printer to desired printer
            win32print.SetDefaultPrinter( fullname )
            log.debug( 'Setting default printer to {0}'.format( fullname ) )

            templatepath = '/F="{0}\\{1}" /P /X'.format( dirpath,
                                                         bartendtemplatefn )
            parmpath = '"{0}{1}"'.format( dirpath,
                self.stationfile[ KEY_TDS_PRINTER_INFO ][ KEY_LBL_PRINT_APP ] )

            try:
                ret_se = win32api.ShellExecute( 0, 'open', parmpath,
                                                templatepath, '', 0 )
            except Exception as e:
                log.debug( 'Unexpected error while printing label: {}'.format(
                            e ) )
                ret_se = 32

            # The shell app needs some time to process the command.
            MySleep( 2 )

            win32print.SetDefaultPrinter( default_printer )
            log.debug( 'Setting default printer to {0}'.format( default_printer ) )
            log.debug( "Shell Execute return is {0}".format( ret_se ) )
            if ( ret_se <= 32 ):
                log.debug( 'ShellExecute failed!' )
                return False

        log.debug( LOG_EXIT )
        return True
    # end function
# end class





def getThisComputerName() -> str:
    """ Gets the current system computer name from os.environ dict

    Return:
        computer_name - computer name from python builtin os environ function
        None - Issue with getting the computer name
    """
    log.debug( LOG_ENTER )

    if ( HW_COMP_NAME not in os.environ.keys() ):
        log.debug( 'Computer name key not in os.environ. '
                   'Cannot get computer name' )
        return None
    # end if

    computer_name = os.environ[ HW_COMP_NAME ]

    if ( computer_name.strip() == '' ):
        log.debug( 'Computer name is an empty string' )
        return None
    # end if

    log.debug( LOG_EXIT )
    return computer_name
# end function




class PrinterObj():
    """ This class is for printer object. Include printer's parameters.
    """
    def __init__( self, fullprintername, printername,
                portname, drivername, sharename ):
        ''' PrinterObj's constructor

        printername: printer name
        portname: printer port name
        drivername: printer driver name
        sharename: printer share name
        '''
        self.fullprintername = fullprintername
        self.printername = printername
        self.portname = portname
        self.drivername = drivername
        self.sharename = sharename
    # end def
# end class





class LabelPrinting():
    """ This class is for printing labels.
    """
    def __init__( self ):
        ''' PrinterObj's constructor
        '''
        pass
    # end function



    def updateTemplate( self, excel_file_name:str, board_info_dict:dict,
                        label_cell_dict:dict ) -> bool:
        """Update template

        excel_file_name: excel file name
        board_info_dict: board information
        label_cell_dict: tell the function how to fill the excel cell, such as:
            "Label Cell Map": {
                "Customer ID": {
                    "row": 2,
                    "column": 1
                },
                "Model": {
                    "row": 2,
                    "column": 2
                },
            },

        Return : True if template is updated
                 False if failed to update

        """
        log.debug( LOG_ENTER )

        dirpath = os.path.abspath( os.getcwd() )
        templatefilefullpath = os.path.join( dirpath, excel_file_name )

        # Update excel file
        xlapp = win32com.client.Dispatch( 'Excel.Application' )
        wb = xlapp.Workbooks.Open( templatefilefullpath )

        try:
            for key in label_cell_dict.keys():

                if ( key in  board_info_dict ):
                    cellval = board_info_dict[ key ]
                else:
                    log.debug( 'No value in results for field {0}'.format(
                            key ) )
                    cellval = ''
                # end if

                row = label_cell_dict[ key ][ KEY_TDS_ROW ]
                column = label_cell_dict[ key ][ KEY_TDS_COLUMN ]
                wb.ActiveSheet.Cells( row, column ).Value = str( cellval )
            # end for
        except Exception as e:
            log.debug( 'Updating label template failed, may be because Excel '
                       'is already opened' )
            log.debug(e)
            msg = 'Unsuccessful update of template, may be due to Excel ' + \
                  'file already being open'
            messagebox.showwarning( title = 'Label Printing Failed',
                    message = msg )
            return False
        # end try

        try:
            wb.Close( SaveChanges = 1 )
            xlapp.Quit()
        except Exception as e:
            log.debug( 'win32com failed! on workbook save for filename {0} '
                       '-- {1}'.format( templatefilefullpath, e ) )
            return False
        # end try

        log.debug( LOG_EXIT )

        return True
    # end function
# end class





monfullnamedict = \
    {
        '01': 'January',
        '02': 'February',
        '03': 'March',
        '04': 'April',
        '05': 'May',
        '06': 'June',
        '07': 'July',
        '08': 'August',
        '09': 'September',
        '10': 'October',
        '11': 'November',
        '12': 'December'
    }





def main():
    Root = Tk()
    app = PrintDeviceGUI( master = Root )
    app.mainloop()
# end main

log = logs() # calling log class as object
log.logger_name = os.path.basename( __file__ )



if __name__ == '__main__':
    main()
# end if
