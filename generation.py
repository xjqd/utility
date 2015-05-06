# vim: tabstop=4 shiftwidth=4 softtabstop=4
# create:    09-25-2014
# tool to generate one excel based on GUI input

import sys
import os
import traceback
import string
import random
from  datetime  import  *
import time
import netaddr
import pprint
from constant import FI_Excel_Constant
from xlwt import Workbook
from vivi.vivi_core.common import vnf_log
from vivi.iaas.openstack.api.nova import flavor_get_by_name

RANDOM_LEN = 255

class RowException(Exception):
    def __init__(self, msg="exception reason"):
        self.reason = msg

    def __str__(self):
        return self.reason


class EntryObject(object):

    def __init__(self, logger=None, vnf_manager=None):
        # register logger
        vnf_log.class_setup_vnf_logger(self, logger)
        self.vnf_manager = vnf_manager

    def ValidateSingleRowInput(self, sheet_name=None, rowNum=None, colNum=None):
        ValidInput = True
        try:
            if not sheet_name:
                raise (RowException("unknown sheet name in Excel file"))
            if not rowNum:
                raise (RowException("unknown to write line number"))
        except:
            ValidInput = False
        finally:
            return ValidInput

    def ValidateMergedRowInput(self, sheet_name=None, startRow=None, endRow=None,
                               startCol=None, endCol=None):
        ValidInput = True
        try:
            if not sheet_name:
                raise (RowException("unknown sheet name in Excel file"))
            if not (startRow and endRow):
                raise (RowException("unknown to write line number"))
        except:
            ValidInput = False
        finally:
            return ValidInput

    def WriteSignleColSignleRow(self, excelMgr=None, sheet_name=None, rowNum=None,
                                colNum=None, value=None, style=None):
        if self.ValidateSingleRowInput(sheet_name, rowNum, colNum):
            sheet_name.write(rowNum, colNum, value, style)
            excelMgr.lineNo += 1

    def WriteMultiColSignleRow(self, excelMgr=None, sheet_name=None, rowNum=None,
                               cols=None, values=None, style=None):
        if self.ValidateSingleRowInput(sheet_name, rowNum, cols):
            idx = 0
            for value in values:
                sheet_name.write(rowNum, idx, str(value), style)
                idx += 1
            excelMgr.lineNo += 1

    def WriteMergedColSingleRow(self, excelMgr=None, sheet_name=None, startRow=None, endRow=None,
                                startCol=None, endCol=None, value=None, style=None):
        if self.ValidateMergedRowInput(sheet_name, startRow, endRow, startCol, endCol):
            sheet_name.write_merge(startRow, endRow, startCol, endCol, value, style)
            excelMgr.lineNo += 1

    def WriteHeader(self, excelMgr=None):
        self.logger.info("---Header--- %s" %self.section.header)
        self.WriteMergedColSingleRow(excelMgr, excelMgr.filename,
                                     excelMgr.lineNo, excelMgr.lineNo,
                                     0, 1,
                                     value=self.section.header, style=self.section.style_header)

    def WriteNote(self, excelMgr=None):
        self.WriteMergedColSingleRow(excelMgr, excelMgr.filename,
                                     excelMgr.lineNo, excelMgr.lineNo,
                                     0, 1,
                                     value=self.section.note, style=self.section.style_note)

    def WriteSubtitle(self, excelMgr=None):
        self.WriteMultiColSignleRow(excelMgr, excelMgr.filename,
                                    excelMgr.lineNo,
                                    len(self.section.subtitle), self.section.subtitle,
                                    self.section.style_para)

    def WriteExcelSectionValue(self, excelMgr=None):
        pass


class ExcelMgr(object):

    def __init__(self, gui_data=None, logger=None, vnf_manager=None):
        # initialize structure for sections owned by object
        self.table = []
        # register logger
        vnf_log.class_setup_vnf_logger(self, logger)
        fi_excel_input = pprint.pformat(gui_data, indent=4, width=80, depth=None)
        self.logger.info("FI Excel received data from GUI:")
        #self.logger.info("\n%s" %fi_excel_input)
        self.logger.info("\n%s" % gui_data)
        # filename and lineNo
        self.filename = None
        self.lineNo = 0
        # List all sections here
        self.basicSection = BasicInfoSection(logger=logger, gui_data=gui_data)
        self.hwSection = HardWareSection(logger=logger, gui_data=gui_data)
        self.vmInfoSection = VMInfoSection(logger=logger, gui_data=gui_data, vnf_manager=vnf_manager)
        self.svcLocSection = SvcLocSection(logger=logger, gui_data=gui_data)
        self.diskSection = DiskSection(logger=logger, gui_data=gui_data)
        self.v4section = V4Section(logger=logger, gui_data=gui_data)
        self.hubsshsolSection = HUBSSHSOLSec(logger=logger, gui_data=gui_data)
        self.hubportspeedSection = HUBPortSpeedSec(logger=logger, gui_data=gui_data)
        self.ipv4intsubnetSection = IPv4IntSubnetSec(logger=logger, gui_data=gui_data)
        self.resipv4addrsection = ResIPv4AddrSec(logger=logger, gui_data=gui_data)
        self.svcIPInfoSection = SvcIPV4InfoSection(logger=logger, gui_data=gui_data)
        self.ipv4defGateway = IPV4DefGatewaySection(logger=logger, gui_data=gui_data)
        self.ipv4staticroute = IPV4StaticRouteSection(logger=logger, gui_data=gui_data)

    def generate_sections(self):
        # The section ordering should be consistent with the excel template
        self.basicSection.GenerateSelfSection()
        self.basicSection.AddToExcelSections(excel_mgr=self)
        self.hwSection.GenerateSelfSection()
        self.hwSection.AddToExcelSections(excel_mgr=self)
        self.vmInfoSection.GenerateSelfSection()
        self.vmInfoSection.AddToExcelSections(excel_mgr=self)
        self.svcLocSection.GenerateSelfSection()
        self.svcLocSection.AddToExcelSections(excel_mgr=self)
        self.diskSection.GenerateSelfSection()
        self.diskSection.AddToExcelSections(excel_mgr=self)
        self.hubsshsolSection.GenerateSelfSection()
        self.hubsshsolSection.AddToExcelSections(excel_mgr=self)
        self.hubportspeedSection.GenerateSelfSection()
        self.hubportspeedSection.AddToExcelSections(excel_mgr=self)
        self.ipv4intsubnetSection.GenerateSelfSection()
        self.ipv4intsubnetSection.AddToExcelSections(excel_mgr=self)
        self.v4section.GenerateSelfSection()
        self.v4section.AddToExcelSections(excel_mgr=self)
        self.resipv4addrsection.GenerateSelfSection()
        self.resipv4addrsection.AddToExcelSections(excel_mgr=self)
        self.svcIPInfoSection.GenerateSelfSection()
        self.svcIPInfoSection.AddToExcelSections(excel_mgr=self)
        self.ipv4defGateway.GenerateSelfSection()
        self.ipv4defGateway.AddToExcelSections(excel_mgr=self)
        self.ipv4staticroute.GenerateSelfSection(self.v4section)
        self.ipv4staticroute.AddToExcelSections(excel_mgr=self)

    def create_excel(self, excel_path=None, sheet_name=None):

        if not self.table:
            sys.stderr.write("error\n")

        # create EXCEL
        wb = Workbook(encoding='utf-8')
        # wb = Workbook()
        sheet = wb.add_sheet(sheet_name, cell_overwrite_ok=True)

        self.filename = sheet

        # Excel Header and update
        sheet.write_merge(self.lineNo, self.lineNo, 0, 1,
                          FI_Excel_Constant.sheetTitle, FI_Excel_Constant.style_header)
        self.lineNo += 1
        sheet.write_merge(self.lineNo, self.lineNo, 0, 1,
                          (FI_Excel_Constant.sheetUpdate +
                           str(datetime.fromtimestamp(time.time()))),
                          FI_Excel_Constant.style_update)
        self.lineNo += 1

        # Set the Excel columns to a fix width
        sheet.col(0).width = 16000
        sheet.col(1).width = 16000

        self.logger.info("start to create FI Excel ... ")
        #for section in ExcelMgr.sections:
        for section in self.table:
            generated = False
            for ele in section:
                if not generated:
                    ele.WriteHeader(self)
                    ele.WriteNote(self)
                    ele.WriteSubtitle(self)
                    generated = True
                ele.WriteExcelSectionValue(self)
        wb.save(excel_path)
        self.logger.info("FI Excel create completely")


class BasicInfo(EntryObject):

    def __init__(self, logger=None, vnf_manager=None, basicSection=None, colName=None, Value=None):
        # register logger
        # vnf_log.class_setup_vnf_logger(self, logger)
        super(BasicInfo, self).__init__(logger=logger, vnf_manager=vnf_manager)
        self.section = basicSection
        self.colName = colName
        self.value = Value

    def WriteExcelSectionValue(self, excelMgr=None):
        para = [self.colName, self.value]
        self.WriteMultiColSignleRow(excelMgr, excelMgr.filename,
                                    excelMgr.lineNo, len(para), para, self.section.style_data)


class SectionBase(object):

    def __init__(self, gui_data=None, logger=None):
        # register logger
        vnf_log.class_setup_vnf_logger(self, logger)
        self.gui_data = gui_data
        self.style_header = FI_Excel_Constant.style_header
        self.style_note = FI_Excel_Constant.style_note
        self.style_para = FI_Excel_Constant.style_para
        self.style_data = FI_Excel_Constant.style_data


class BasicInfoSection(SectionBase):

    def __init__(self, gui_data=None, logger=None):
        super(BasicInfoSection, self).__init__(gui_data=gui_data, logger=logger)
        self.basiclist = []
        self.header = FI_Excel_Constant.BasicInfoHeader
        self.note = FI_Excel_Constant.BasicInfoNote
        self.subtitle = ["# Parameter Name", "Parameter Values"]

    def FormatBasicEntry(self, colName=None, Value=None):
        subnet = BasicInfo(colName=colName, Value=Value, logger=self.logger, basicSection=self)
        return subnet

    def GenerateSelfSection(self):
        self.logger.info("---Start Generating BasicInfoSection---")
        for col in FI_Excel_Constant.BasicCols:
            if col == "Chassis Type":
                value = "Virtualized System"
            elif col == "Tenant Type":
                value = "virtual"
            elif col == "Network Type":
                value = "legacy"
            elif col == "SNMP Community":
                value = "public"
            elif col == "System Name":
                # use the system_prefix value fill vnf_system
                value = self.gui_data["vnf"]["system_prefix"]
            elif col == "System Prefix":
                value = self.gui_data["vnf"]["system_prefix"]
            elif col == "Application Name":
                value = ""
                if self.gui_data["vnf"]["vnf_application"]:
                    for val in self.gui_data["vnf"]["vnf_application"]:
                        value += val + ", "
                    value = value.rstrip(", ")
                else:
                    value = ""
            elif col == "Time Zone":
                value = self.gui_data["vnf"]["vnf_adv_setting"]["time_zone"]
            elif col == "Use SBPR (Source Based Policy Routing)":
                value = self.gui_data["vnf"]["vnf_adv_setting"]["use_sbpr"]
            elif col == "Local DNS Domain":
                value = self.gui_data["vnf"]["vnf_adv_setting"]["local_dns_domain"]
            elif col == "Local ENUM Domain":
                value = self.gui_data["vnf"]["vnf_adv_setting"]["local_enum_domain"]
            elif col == "Remote DNS Domain":
                value = self.gui_data["vnf"]["vnf_adv_setting"]["remote_dns_domain"]
            elif col == "Remote ENUM Domain":
                value = self.gui_data["vnf"]["vnf_adv_setting"]["remote_enum_domain"]
            elif col == "NTP Server IPv4 Address":
                value = self.gui_data["vnf"]["vnf_adv_setting"]["ntp_ipv4"]
            elif col == "NTP Server IPv6 Address":
                value = self.gui_data["vnf"]["vnf_adv_setting"]["ntp_ipv6"]
            elif col == "DNS Server IPv4 Address/DNS Domain Name":
                value = self.gui_data["vnf"]["vnf_adv_setting"]["dns_ipv4_ip_and_domain"]
            elif col == "ENUM Server IPv4 Address/ENUM Domain Name":
                value = self.gui_data["vnf"]["vnf_adv_setting"]["enum_ipv4_ip_and_domain"]
            elif col == "DNS Server IPv6 Address/DNS Domain Name":
                value = self.gui_data["vnf"]["vnf_adv_setting"]["dns_ipv6_ip_and_domain"]
            elif col == "License Reference Name string":
                value = self.gui_data["vnf"]["license_reference_name"]
                if not value:
                    value = ''.join(random.SystemRandom().choice(string.uppercase + string.digits) for _ in xrange(RANDOM_LEN))
            else:
                value = self.gui_data["vnf"]["vnf_adv_setting"]["enum_ipv6_ip_and_domain"]
            basic = self.FormatBasicEntry(col, value)
            self.basiclist.append(basic)
        self.logger.info("---End Generating BasicInfoSection---")

    def AddToExcelSections(self, excel_mgr=None):
        self.logger.info("---start to Add BasicInfoSection into list ---")
        #ExcelMgr.sections.append(self.basiclist)
        excel_mgr.table.append(self.basiclist)
        self.logger.info("---end to Add BasicInfoSection into list ---")


class HardWare(EntryObject):

    def __init__(self, logger=None, vnf_manager=None, hwSection=None, ColName=None, Value=None):
        # register logger
        # vnf_log.class_setup_vnf_logger(self, logger)
        super(HardWare, self).__init__(logger=logger, vnf_manager=vnf_manager)
        self.section = hwSection
        self.colName = ColName
        self.Value = Value

    def WriteExcelSectionValue(self, excelMgr=None):
        para = [self.colName, self.Value]
        self.WriteMultiColSignleRow(excelMgr, excelMgr.filename,
                                    excelMgr.lineNo, len(para), para, self.section.style_data)


class HardWareSection(SectionBase):

    def __init__(self, gui_data=None, logger=None):
        super(HardWareSection, self).__init__(gui_data=gui_data, logger=logger)
        self.hardwarelist = []
        self.header = FI_Excel_Constant.HwHeader
        self.note = FI_Excel_Constant.HwNote
        self.subtitle = ["# Parameter Name", "Parameter Values"]

    def FormatHardWare(self, colName=None, Value=None):
        hardware = HardWare(ColName=colName, Value=Value, logger=self.logger, hwSection=self)
        return hardware

    def GenerateSelfSection(self):
        vnfc_list = []
        Excel_HwValue = ""
        for vnfcId in sorted(self.gui_data["vnfc"].keys()):
            vnfc_list.append(int(vnfcId))
        for vnfc_id in vnfc_list:
            Excel_HwValue += "0/"+ str(vnfc_id) + "/_GUEST/NONE" + "\n"
        for vnfc_id in vnfc_list:
            Excel_HwValue += "1/"+ str(vnfc_id) + "/_GUEST/NONE" + "\n"
        for col in FI_Excel_Constant.Excel_Hardware:
            if col == "Cabinet":
                value = "0"
            elif col == "Cabinet/Shelf/Mnemonic":
                value = "0/0/_SHELF_1_99,0/1/_SHELF_1_99"
            else:
                # value = FI_Excel_Constant.Excel_HwValue
                value = Excel_HwValue
            hardware = self.FormatHardWare(col, value)
            self.hardwarelist.append(hardware)

    def AddToExcelSections(self, excel_mgr=None):
        excel_mgr.table.append(self.hardwarelist)


class IPV4SubNet(EntryObject):

    def __init__(self, logger=None, vnf_manager=None,
                 v4section=None, name=None, lcp_subnet_id=None,
                 net_id=None, subnet_id=None, Base=None, Mask=None, Gateway=None,
                 Transport=None, Redundant=None, Vlantag=None, ConnectIP=None):
        # register logger
        # vnf_log.class_setup_vnf_logger(self, logger)
        super(IPV4SubNet, self).__init__(logger=logger, vnf_manager=vnf_manager)
        self.section = v4section
        self.lcp_subnet_id = lcp_subnet_id
        self.net_id = net_id
        self.subnet_id = subnet_id
        self.name = name
        self.base = Base
        self.mask = Mask
        self.gateway = Gateway
        self.transport = Transport
        self.redundant = Redundant
        self.vlantag = Vlantag
        self.connectip = ConnectIP

    def WriteSubtitle(self, excelMgr=None):
        pass

    def WriteExcelSectionValue(self, excelMgr=None):
        fixedStr = "IPv4 Subnet %s " % (str(self.lcp_subnet_id))
        for rowname in FI_Excel_Constant.Excel_V4Subnet:
            if rowname == 'Name':
                para = [fixedStr + rowname, self.name]
            elif rowname == 'Base':
                para = [fixedStr + rowname, self.base]
            elif rowname == 'Network Mask':
                para = [fixedStr + rowname, self.mask]
            elif rowname == 'Gateway':
                para = [fixedStr + rowname, self.gateway]
            elif rowname == 'Transport Interface':
                para = [fixedStr + rowname, self.transport]
            elif rowname == 'Redundancy Mode':
                para = [fixedStr + rowname, self.redundant]
            elif rowname == 'VLAN Tag':
                para = [fixedStr + rowname, self.vlantag]
            else:
                para = [fixedStr + rowname, self.connectip]
            self.WriteMultiColSignleRow(excelMgr, excelMgr.filename,
                                        excelMgr.lineNo, len(para), para, self.section.style_data)


class V4Section(SectionBase):

    def __init__(self, gui_data=None, logger=None):
        super(V4Section, self).__init__(gui_data=gui_data, logger=logger)
        self.v4list = []
        self.header = FI_Excel_Constant.Ipv4SubHeader
        self.note = FI_Excel_Constant.Ipv4SubNote
        self.subtitle = ["para", "value"]
        self.v4net = {}

    def FormatIPV4SubNet(self, version=None, name=None, Excel_netid=None, subnetVal=None):
        base = subnetVal["netmask"].split("/")[0]
        subnet = IPV4SubNet(name=name,
                            lcp_subnet_id=subnetVal["lcp_subnet_id"],
                            net_id=subnetVal["net_id"],
                            subnet_id=subnetVal["subnet_id"],
                            # Base=subnetVal["start_ip"],
                            Base=base,
                            Mask=str((netaddr.IPNetwork(subnetVal["netmask"])).netmask),
                            Gateway=subnetVal["gateway"],
                            Transport="Just Front Port - " + subnetVal["vif"],
                            Redundant="none",
                            Vlantag="",
                            ConnectIP="",
                            logger=self.logger,
                            v4section=self)
        self.v4net[name] = subnetVal["lcp_subnet_id"]
        return subnet

    def GenerateSelfSection(self):
        for (ipVer, ipVal) in self.gui_data["subnet"].items():
            if ipVer == "IPv4":
                for (subNetName, subNetVal) in ipVal.items():
                    subnet = self.FormatIPV4SubNet(version=subNetVal["ip_version"],
                                                   name=subNetName,
                                                   subnetVal=subNetVal
                                                   )
                    self.v4list.append(subnet)
            elif ipVer == "IPv6":
                pass

    def AddToExcelSections(self, excel_mgr=None):
        excel_mgr.table.append(self.v4list)


class HUBSSHSOLInfo(EntryObject):

    def __init__(self, logger=None, vnf_manager=None,
                 hubsshsolsec=None, shelfid=None, hubid=None,
                 sship="", sshmask="", sshgateway="", solip=""):
        # register logger
        # vnf_log.class_setup_vnf_logger(self, logger)
        super(HUBSSHSOLInfo, self).__init__(logger=logger, vnf_manager=vnf_manager)
        self.section = hubsshsolsec
        self.shelfId = shelfid
        self.hubId = hubid
        self.sshIp = sship
        self.sshMask = sshmask
        self.sshGateway = sshgateway
        self.solIp = solip

    def WriteSubtitle(self, excelMgr=None):
        pass

    def WriteExcelSectionValue(self, excelMgr=None):
        NetRowlist = ["SSH IP Address", "SSH Netmask", "SSH Gateway", "SOL IP Address"]
        for sid in self.shelfId:
            for hid in self.hubId:
                for rowname in NetRowlist:
                    fixedStr = "Shelf-%s Hub-%s " % (str(sid), str(hid))
                    if rowname == 'SSH IP Address':
                        para = [fixedStr + rowname, self.sshIp]
                    elif rowname == 'SSH Netmask':
                        para = [fixedStr + rowname, self.sshMask]
                    elif rowname == 'SSH Gateway':
                        para = [fixedStr + rowname, self.sshGateway]
                    elif rowname == 'SOL IP Address':
                        para = [fixedStr + rowname, self.solIp]
                    self.WriteMultiColSignleRow(excelMgr, excelMgr.filename, excelMgr.lineNo,
                                                len(para), para, self.section.style_data)


class HUBSSHSOLSec(SectionBase):

    def __init__(self, gui_data=None, logger=None):
        super(HUBSSHSOLSec, self).__init__(gui_data=gui_data, logger=logger)
        self.hubsshlist = []
        self.header = FI_Excel_Constant.HUBSSHSOLHeader
        self.note = FI_Excel_Constant.HUBSSHSOLNote
        self.subtitle = ["Address", "addr"]

    def FormatHUBSSHSOLInfo(self, shelfid, hubid, sship, sshmask, sshgateway, solip):
        subnet = HUBSSHSOLInfo(shelfid=shelfid, hubid=hubid, sship=sship,
                               sshmask=sshmask, sshgateway=sshgateway, solip=solip,
                               logger=self.logger, hubsshsolsec=self)
        return subnet

    def GenerateSelfSection(self):
        subnet = self.FormatHUBSSHSOLInfo([0, 1], [7, 8], "", "", "", "")
        self.hubsshlist.append(subnet)

    def AddToExcelSections(self, excel_mgr=None):
        excel_mgr.table.append(self.hubsshlist)


class HUBPortSpeedInfo(EntryObject):

    def __init__(self, logger=None, vnf_manager=None,
                 hubportspeedsec=None, hubid=None, fgeid=None, rgeid=None):
        # register logger
        # vnf_log.class_setup_vnf_logger(self, logger)
        super(HUBPortSpeedInfo, self).__init__(logger=logger, vnf_manager=vnf_manager)
        self.section = hubportspeedsec
        self.hubId = hubid
        self.fgeId = fgeid
        self.rgeId = rgeid

    def WriteExcelSectionValue(self, excelMgr=None):
        for hid in self.hubId:
            for gid in self.fgeId:
                para = ["Hub-%s Front GE%s" % (str(hid), str(gid)), "Auto"]
                self.WriteMultiColSignleRow(excelMgr, excelMgr.filename, excelMgr.lineNo,
                                            len(para), para, self.section.style_data)
        for hid in self.hubId:
            for gid in self.rgeId:
                para = ["Hub-%s Rear GE%s" % (str(hid), str(gid)), "Auto"]
                self.WriteMultiColSignleRow(excelMgr, excelMgr.filename, excelMgr.lineNo,
                                            len(para), para, self.section.style_data)


class HUBPortSpeedSec(SectionBase):

    def __init__(self, gui_data=None, logger=None):
        super(HUBPortSpeedSec, self).__init__(gui_data=gui_data, logger=logger)
        self.hubportlist = []
        self.header = FI_Excel_Constant.HUBPortSpeedHeader
        self.note = FI_Excel_Constant.HUBPortSpeedNote
        self.subtitle = ["# Port Name", "Speed"]

    def FormatHUBPortSpeedInfo(self, hubid, fgeid, rgeid):
        subnet = HUBPortSpeedInfo(hubid=hubid, fgeid=fgeid, rgeid=rgeid,
                                  logger=self.logger, hubportspeedsec=self)
        return subnet

    def GenerateSelfSection(self):
        subnet = self.FormatHUBPortSpeedInfo([7, 8], [0, 1, 3], [0, 1, 2, 3, 4, 5, 6, 7])
        self.hubportlist.append(subnet)

    def AddToExcelSections(self, excel_mgr=None):
        excel_mgr.table.append(self.hubportlist)


class IPv4IntSubnetInfo(EntryObject):

    def __init__(self, logger=None, vnf_manager=None, ipv4intsubsec=None):
        # register logger
        # vnf_log.class_setup_vnf_logger(self, logger)
        super(IPv4IntSubnetInfo, self).__init__(logger=logger, vnf_manager=vnf_manager)
        self.section = ipv4intsubsec

    def WriteSubtitle(self, excelMgr=None):
        pass

    def WriteExcelSectionValue(self, excelMgr=None):
        NetRowlist = ["Base", "Transport Interface", "Redundancy Mode", "VLAN Tag"]
        fixedStr = "IPv4 Internal Subnet 0 "
        for rowname in NetRowlist:
            if rowname == 'Base':
                para = [fixedStr + rowname, "169.254.0.0"]
            if rowname == 'Transport Interface':
                para = [fixedStr + rowname, "Both Rear Port - eth0 and eth1"]
            if rowname == 'Redundancy Mode':
                para = [fixedStr + rowname, "iipm"]
            if rowname == 'VLAN Tag':
                para = [fixedStr + rowname, ""]
            self.WriteMultiColSignleRow(excelMgr, excelMgr.filename, excelMgr.lineNo,
                                        len(para), para, self.section.style_data)


class IPv4IntSubnetSec(SectionBase):

    def __init__(self, gui_data=None, logger=None):
        super(IPv4IntSubnetSec, self).__init__(gui_data=gui_data, logger=logger)
        self.ipv4subnetlist = []
        self.header = FI_Excel_Constant.IPv4IntSubnetHeader
        self.note = FI_Excel_Constant.IPv4IntSubnetNote

    def FormatIPv4IntSubnetInfo(self):
        subnet = IPv4IntSubnetInfo(logger=self.logger, ipv4intsubsec=self)
        return subnet

    def GenerateSelfSection(self):
        subnet = self.FormatIPv4IntSubnetInfo()
        self.ipv4subnetlist.append(subnet)

    def AddToExcelSections(self, excel_mgr=None):
        excel_mgr.table.append(self.ipv4subnetlist)


class ResIPv4AddrInfo(EntryObject):

    def __init__(self, logger=None, vnf_manager=None, resIPv4Addrsec=None):
        # register logger
        # vnf_log.class_setup_vnf_logger(self, logger)
        super(ResIPv4AddrInfo, self).__init__(logger=logger, vnf_manager=vnf_manager)
        self.section = resIPv4Addrsec

    def WriteExcelSectionValue(self, excelMgr=None):
        para = ["Reserved IPv4 Address", ""]
        self.WriteMultiColSignleRow(excelMgr, excelMgr.filename, excelMgr.lineNo,
                                    len(para), para, self.section.style_data)


class ResIPv4AddrSec(SectionBase):

    def __init__(self, gui_data=None, logger=None):
        super(ResIPv4AddrSec, self).__init__(gui_data=gui_data, logger=logger)
        self.resipv4addrlist = []
        self.header = FI_Excel_Constant.ResIPv4AddrHeader
        self.note = FI_Excel_Constant.ResIPv4AddrNote
        self.subtitle = ["# Reserved IPv4 Address Prefix", "Subnet ID/IP Address/Description"]

    def FormatResIPv4AddrInfo(self):
        subnet = ResIPv4AddrInfo(logger=self.logger, resIPv4Addrsec=self)
        return subnet

    def GenerateSelfSection(self):
        subnet = self.FormatResIPv4AddrInfo()
        self.resipv4addrlist.append(subnet)

    def AddToExcelSections(self, excel_mgr=None):
        excel_mgr.table.append(self.resipv4addrlist)


class DiskSchema(EntryObject):

    def __init__(self, logger=None, vnf_manager=None, diskSection=None, vnfcId=None, diskschema=None):
        # register logger
        # vnf_log.class_setup_vnf_logger(self, logger)
        super(DiskSchema, self).__init__(logger=logger, vnf_manager=vnf_manager)
        if not vnfcId:
            # replace it as the offical logger machanism
            self.logger.info("DiskSection invlaid vnfcId value")
            return None
        self.colName = "Disk Scheme:s00c" + vnfcId + "h0," + "s01c" + vnfcId + "h0"
        self.diskschema = diskschema
        self.logger = logger
        self.section = diskSection

    def WriteHeader(self, excelMgr=None):
        pass

    def WriteNote(self, excelMgr=None):
        pass

    def WriteExcelSectionValue(self, excelMgr=None):
            para = [self.colName, self.diskschema]
            self.WriteMultiColSignleRow(excelMgr, excelMgr.filename,
                                        excelMgr.lineNo, len(para), para, self.section.style_data)


class DiskSection(SectionBase):

    def __init__(self, gui_data=None, logger=None):
        super(DiskSection, self).__init__(gui_data=gui_data, logger=logger)
        self.disklist = []
        self.subtitle = ["# Disk Scheme:[shelf-slot-host]", "diskid/scheme"]

    def FormatDiskShema(self, vnfcId=None, schema=None):
        schema = DiskSchema(vnfcId=vnfcId, diskschema=schema, logger=self.logger, diskSection=self)
        return schema

    def GenerateSelfSection(self):
        for vnfcId in sorted(self.gui_data["vnfc"]):
            diskschema = self.FormatDiskShema(vnfcId=vnfcId, schema="front-1/LCP_vm_oam_127G")
            self.disklist.append(diskschema)

    def AddToExcelSections(self, excel_mgr=None):
        excel_mgr.table.append(self.disklist)


class ServiceLocation(EntryObject):

    def __init__(self, logger=None, vnf_manager=None, svcLocSection=None, vnfcId=None, svcpool=[]):
        # register logger
        # vnf_log.class_setup_vnf_logger(self, logger)
        super(ServiceLocation, self).__init__(logger=logger, vnf_manager=vnf_manager)
        if not vnfcId:
            return None
        self.colName = "Service Location:s00c" + vnfcId + "h0," + "s01c" + vnfcId + "h0"
        self.svcpool = svcpool
        self.logger = logger
        self.section = svcLocSection

    def WriteExcelSectionValue(self, excelMgr=None):
            if type(self.svcpool) is list:
                value = ","
                for val in self.svcpool:
                    value = value + val + ","
            value = value.strip(",")
            para = [self.colName, value]
            self.WriteMultiColSignleRow(excelMgr, excelMgr.filename,
                                        excelMgr.lineNo, len(para), para, self.section.style_data)


class SvcLocSection(SectionBase):

    def __init__(self, gui_data=None, logger=None):
        super(SvcLocSection, self).__init__(gui_data=gui_data, logger=logger)
        self.header = FI_Excel_Constant.SvcLocationHeader
        self.note = FI_Excel_Constant.SvcLocationNote
        self.subtitle = ["# Service Location: [shelf-slot-host]", "service-pool"]
        self.svcLocation = []

    def FormatSvcLocation(self, vnfcId=None, svcpool=None):
        svcloc = ServiceLocation(logger=self.logger, svcLocSection=self,
                                 vnfcId=str(vnfcId), svcpool=svcpool)
        return svcloc

    def GenerateSelfSection(self):
        for vnfcId in sorted(self.gui_data["vnfc"]):
            svcpool = []
            for svc in self.gui_data["vnfc"][vnfcId]["service"]:
                svcpool.append(svc)
            svcLocation = self.FormatSvcLocation(vnfcId=vnfcId, svcpool=svcpool)
            self.svcLocation.append(svcLocation)

    def AddToExcelSections(self, excel_mgr=None):
        excel_mgr.table.append(self.svcLocation)


class VMInfo(EntryObject):

    def __init__(self, logger=None, vnf_manager=None, vmInfoSection=None, vnfcId=None, flavor=None):
        super(VMInfo, self).__init__(logger=logger, vnf_manager=vnf_manager)
        # register logger
        # vnf_log.class_setup_vnf_logger(self, logger)
        self.colName = "VM Name:s00c" + vnfcId + "h0," + "s01c" + vnfcId + "h0"
        self.flavor = flavor
        self.logger = logger
        self.section = vmInfoSection

    def WriteExcelSectionValue(self, excelMgr=None):
            para = [self.colName, self.flavor]
            self.WriteMultiColSignleRow(excelMgr, excelMgr.filename,
                                        excelMgr.lineNo, len(para), para, self.section.style_data)


class VMInfoSection(SectionBase):

    def __init__(self, gui_data=None, logger=None, vnf_manager=None):
        super(VMInfoSection, self).__init__(gui_data=gui_data, logger=logger)
        self.vminfolist = []
        self.header = FI_Excel_Constant.VMHeader
        self.note = FI_Excel_Constant.VMNote
        self.subtitle = ["# VM pair Name", "core/memory/disk_size(in GB)"]
        self.vnf_manager = vnf_manager

    def FormatVMInfo(self, vnfcId=None, flavor=None):
        return VMInfo(logger=self.logger, vmInfoSection=self, vnfcId=vnfcId, flavor=flavor)

    def GenerateSelfSection(self):
        for vnfcId in sorted(self.gui_data["vnfc"].keys()):
            flavor_name = self.gui_data["vnfc"][vnfcId]["flavor"]
            flavors = flavor_get_by_name(self.vnf_manager.context, flavor_name)
            flavor = str(flavors.vcpus) + "/" + str(flavors.ram/1024) + "/" + str(flavors.disk)
            vminfo = self.FormatVMInfo(vnfcId, flavor)
            self.vminfolist.append(vminfo)

    def AddToExcelSections(self, excel_mgr=None):
        excel_mgr.table.append(self.vminfolist)


class SvcIPInfo(EntryObject):

    def __init__(self, logger=None, vnf_manager=None, svcInfoSection=None, svcname=None, subId=None, ipAddr=None):
        # register logger
        # vnf_log.class_setup_vnf_logger(self, logger)
        super(SvcIPInfo, self).__init__(logger=logger, vnf_manager=vnf_manager)
        self.section = svcInfoSection
        self.colName = "IPv4 Service IP: " + svcname
        self.svcip = subId + "/" + ipAddr

    def WriteExcelSectionValue(self, excelMgr=None):
            para = [self.colName, self.svcip]
            self.WriteMultiColSignleRow(excelMgr, excelMgr.filename,
                                        excelMgr.lineNo, len(para), para, self.section.style_data)


class SvcIPV4InfoSection(SectionBase):

    def __init__(self, gui_data=None, logger=None):
        super(SvcIPV4InfoSection, self).__init__(gui_data=gui_data, logger=logger)
        self.header = FI_Excel_Constant.IPv4ServiceIPInfoHeader
        self.note = FI_Excel_Constant.IPv4ServiceIPInfoNote
        self.subtitle = FI_Excel_Constant.IPv4ServiceIPInfoSubtitle
        self.svcipv4infolist = []

    def FormatSvcIPInfo(self, svcname=None, subId=None, ipaddr=None):
        return SvcIPInfo(logger=self.logger, svcInfoSection=self, svcname=svcname,
                         subId=subId, ipAddr=ipaddr)

    def GenerateSelfSection(self):
        for vnfcId in sorted(self.gui_data["vnfc"]):
            for (svc, svcVal) in self.gui_data["vnfc"][vnfcId]["service"].items():
                for (ipVer, niVal) in svcVal["ni"].items():
                    if ipVer == "IPv4":
                        net = self.gui_data["subnet"]["IPv4"]
                        for (niName, val) in niVal.items():
                            subnet = net[val["subnet"]]
                            subId = subnet["lcp_subnet_id"]
                            for ele in val["ni_fixed_floating_type"]:
                                if ele == "fixed":
                                    fixip = val["fixed_ip_list"]
                                    for member in ['0', '1']:
                                        svcname = "fixed," + svc.upper() + "-" + member \
                                                  + "," + niName
                                        svcIPInfo = self.FormatSvcIPInfo(svcname=svcname,
                                                                         subId=str(subId),
                                                                         ipaddr=fixip[int(member)])
                                        self.svcipv4infolist.append(svcIPInfo)
                                elif ele == "floating":
                                    if int(val["floating_ip_num"]) == 0:
                                        continue
                                    fltip = val["floating_ip_list"]
                                    if not fltip:
                                        continue
                                    svcname = "floating," + svc.upper() + "-X" + "," + niName
                                    addr = ""
                                    for ip in fltip:
                                        addr += str(subId) + "/" + ip
                                        addr += ","
                                    addr = addr.split('/', 1)[1]
                                    addr = addr.rstrip(",")
                                    svcIPInfo = self.FormatSvcIPInfo(svcname=svcname,
                                                                     subId=str(subId),
                                                                     ipaddr=addr)
                                    self.svcipv4infolist.append(svcIPInfo)
                    elif ipVer == "IPv6":
                        pass

    def AddToExcelSections(self, excel_mgr=None):
        excel_mgr.table.append(self.svcipv4infolist)


class DefGateway(EntryObject):

    def __init__(self, logger=None, vnf_manager=None, DefGwSection=None, vnfcId=None, subId=None):
        # register logger
        # vnf_log.class_setup_vnf_logger(self, logger)
        super(DefGateway, self).__init__(logger=logger, vnf_manager=vnf_manager)
        self.section = DefGwSection
        self.colName = "Default IPv4 Gateway: " + vnfcId + "h0"
        self.subId = subId

    def WriteExcelSectionValue(self, excelMgr=None):
            para = [self.colName, self.subId]
            self.WriteMultiColSignleRow(excelMgr, excelMgr.filename,
                                        excelMgr.lineNo, len(para), para, self.section.style_data)


class IPV4DefGatewaySection(SectionBase):

    def __init__(self, gui_data=None, logger=None):
        super(IPV4DefGatewaySection, self).__init__(gui_data=gui_data, logger=logger)
        self.header = FI_Excel_Constant.IPv4DefGwHeader
        self.note = FI_Excel_Constant.IPv4DefGwNote
        self.subtitle = FI_Excel_Constant.IPv4DefGwSubtile
        self.ipv4defgwlist = []

    def FormatIPDefGw(self, vnfcId=None, subId=None):
        return DefGateway(logger=self.logger, DefGwSection=self, vnfcId=vnfcId, subId=subId)

    def GenerateSelfSection(self):
        for vnfcId in sorted(self.gui_data["vnfc"]):
            subnetName = self.gui_data["vnfc"][vnfcId]["subnet"]
            subId = self.gui_data["subnet"]["IPv4"][subnetName]["lcp_subnet_id"]
            for val in ["s00c", "s01c"]:
                defaultgw = self.FormatIPDefGw(vnfcId=val+vnfcId,
                                               subId=subId)
                self.ipv4defgwlist.append(defaultgw)

    def AddToExcelSections(self, excel_mgr=None):
        excel_mgr.table.append(self.ipv4defgwlist)


class StaticRoute(EntryObject):

    def __init__(self, logger=None, vnf_manager=None, StaticRouteSection=None,
                idx=None, dest_subnet=None, dest_mask=None, src_subnet=None,
                svc=None):
        # register logger
        vnf_log.class_setup_vnf_logger(self, logger)
        super(StaticRoute, self).__init__(logger=logger, vnf_manager=vnf_manager)
        self.section = StaticRouteSection
        self.idx = idx
        self.dest_subnet = dest_subnet
        self.dest_mask = dest_mask
        self.src_subnet = src_subnet
        self.svc = svc

    def WriteExcelSectionValue(self, excelMgr=None):
        for rowname in FI_Excel_Constant.Excel_V4StaticRoute:
            if rowname == 'Destination Subnet':
                para = [self.idx + rowname, self.dest_subnet]
            elif rowname == 'Destination Subnet Mask':
                para = [self.idx + rowname, self.dest_mask]
            elif rowname == 'External Subnet for Source IP':
                para = [self.idx + rowname, self.src_subnet]
            else:
                para = [self.idx + rowname, self.svc]
            self.WriteMultiColSignleRow(excelMgr, excelMgr.filename,
                                        excelMgr.lineNo, len(para), para, self.section.style_data)

    def WriteSubtitle(self, excelMgr=None):
        pass

class IPV4StaticRouteSection(SectionBase):

    def __init__(self, gui_data=None, logger=None):
        super(IPV4StaticRouteSection, self).__init__(gui_data=gui_data, logger=logger)
        self.header = FI_Excel_Constant.IPv4StaticRoute
        self.note = FI_Excel_Constant.IPv4StaticRouteNote
        self.ipv4staticroutelist = []
        self.netv4pool = None

    def find_netId_by_name(self, name=None):
        if not self.netv4pool.v4net:
            self.logger.info("No value saved in v4net")
        for key, val in self.netv4pool.v4net.items():
            if key == name:
                self.logger.info("val=%s", val)
                return val

    def FormatStaticRoute(self, routeIdx=None, dest=None, src=None, svc=None):
        net_obj = netaddr.IPNetwork(dest)
        subnet = str(net_obj.ip)
        mask = str(net_obj.netmask)
        return StaticRoute(logger=self.logger, StaticRouteSection=self, idx=routeIdx,
                           dest_subnet=subnet, dest_mask=mask, src_subnet=src, svc=svc)

    def GenerateSelfSection(self, v4subnetpool=None):
        self.netv4pool = v4subnetpool
        for key, val in v4subnetpool.v4net.items():
            self.logger.info("init name=%s, val=%s" %(key, val))
        for name, values in self.gui_data["static_route"].items():
            if name.upper() == "IPV4":
                idx = 1
                for val in values:
                    destination = val['destination_cidr']
                    src_name = val['source_subnet']
                    src_subnet = self.find_netId_by_name(src_name)
                    svc = val['service']
                    staticroute= self.FormatStaticRoute(routeIdx="IPv4 Static Route " + str(idx) + ": ",
                                                        dest=destination,
                                                        src=src_subnet,
                                                        svc=svc)
                    self.ipv4staticroutelist.append(staticroute)
                    idx += 1
            elif name.upper() == "IPV6":
                pass

    def AddToExcelSections(self, excel_mgr=None):
        excel_mgr.table.append(self.ipv4staticroutelist)

if __name__ == '__main__':
    try:
        logdata = LogData()
        logger = LogMgr(logdata)
        sections = ExcelMgr(logger=None, gui_data=FI_Excel_Constant.deployment4)
        sections.generate_sections()
        sections.create_excel(excel_path="lcp.xls", sheet_name="ics")
    except (RowException, Exception), ex:
        traceback.print_exc(file=logger.logfile)
        logger.logfile.write("\n")
        logger.CriticalLog(section="Exception", reason=str(ex))
    finally:
        logger.SaveLog()
