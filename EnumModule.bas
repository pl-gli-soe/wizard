Attribute VB_Name = "EnumModule"

' zbior enumow przenosze do osobnego modulu

' wraz z niedziela!
' 2015-10-11
' jest duplikat na enumach
' wszystkie kolumny + kolumny ktore posiadaja formule
' wskazuja wartosciowo dokladnie to samo
' jednak celem dla formul bylo uszczuplenie listy wyboru, gdy skupiamy sie na formula
' handling

Public Enum E_IS_JENNY
    E_JENNY_JENNY
    E_JENNY_NO_JENNY
End Enum

Public Enum E_TYPE_OF_COMMENT
    E_PONO_CMNT = 1
    E_ONE_LINE_CMNT
End Enum


Public Enum E_PIVOTS
    E_PIVOT_RESP = 1
    E_PIVOT_PPAP_STATUS
    E_PIVOT_COUNTRY_CODE
    E_PIVOT_FUP_CODE
    E_PIVOT_INTRANSIT_MRD
    E_PIVOT_INTRANSIT_TODAY
    E_PIVOT_INTRANSIT_CUSTOM_DATE
End Enum

Public Enum E_ADD_EDIT_PUSES
    E_ADD = 1
    E_EDIT = 2
End Enum

Public Enum E_COLUMNS_WITH_FORMULAS
    E_F_TOTAL = 16
    E_F_MRD1_ORDERED_STATUS = 20
    E_F_MRD1_CONFIRMED_STATUS = 22
    E_F_MRD1_PUS_STATUS = 23
    E_F_MRD2_ORDERED_STATUS = 26
    E_F_MRD2_CONFIRMED_STATUS = 28
    E_F_MRD2_PUS_STATUS = 29
    E_F_TOTAL_PUS = 33
    E_F_TOTAL_PUS_STATUS = 34
End Enum



Public Enum E_ADD_DATA
    E_DOPISZ
    E_NADPISZ
End Enum

Public Enum E_DATE_OR_CW
    E_DC_DATE
    E_DC_CW
End Enum

Public Enum E_DETAILS_WIZARD_ORDER
    PIERWSZY
    SRODEK
    ostatni
End Enum



Public Enum E_JAKI_FORM
    CFG = 1
    WIZARD_COMBOBOX
    WIZARD_DATEPICKER
    WIZARD_TOGGLE
    WIZARD_TXTBOX
End Enum


Public Enum E_PUS_SH
    O_INDX = 1
    O_PN
    O_DUNS
    O_FUP_code
    O_Pick_up_date
    O_Delivery_Date
    O_Pick_up_Qty
    O_PUS_Number
End Enum

Public Enum E_NEW_PROJECT_ITEM
    PLT = 1
    PROJECT
    BIW_GA ' BIW or GA
    MY
    PHAZE
    BOM
    PICKUP_DATE
    PPAP_GATE
    mrd
    BUILD_START
    BUILD_END
    KOORDYNATOR
    E_ACTIVE
    CAPACITY_CHECK
    E_MRD_DATE
    E_MRD_REG_ROUTES
    E_PLATFORM
    E_TRANSPORTATION_ACCOUNT_NUMBER
    E_UNIQUE_ID
End Enum


Public Enum E_MASTER_MANDATORY_COLUMNS
    pn = 1
    Alternative_PN
    PN_Name
    GPDS_PN_Name
    duns
    Supplier_Name
    country_code
    MGO_code
    Responsibility
    fup_code
    SQ
    ppap_status
    SQ_Comments
    MRD1_QTY
    MRD2_QTY
    Total_QTY
    ADD_to_T_slash_D
    MRD1_Ordered_date
    MRD1_Ordered_QTY
    MRD1_Ordered_STATUS
    MRD1_confirmed_qty
    MRD1_confirmed_qty_dot__Status
    MRD1_Total_PUS_STATUS
    MRD2_Ordered_date
    MRD2_Ordered_QTY
    MRD2_Ordered_STATUS
    MRD2_confirmed_qty
    MRD2_confirmed_qty_dot__Status
    MRD2_Total_PUS_STATUS
    Delivery_confirmation
    First_Confirmed_PUS_Date
    Delivery_reconfirmation
    Total_PUS_QTY
    Total_PUS_STATUS
    Comments
    Bottleneck
    Future_Osea
    DRE
    EDI_Received
    Capacity
    Oncost_confirmation
    BLANK3
    BLANK4
End Enum


