<?xml version='1.0' encoding='UTF-8'?>
<Project Type="Project" LVVersion="14008000">
	<Item Name="我的电脑" Type="My Computer">
		<Property Name="server.app.propertiesEnabled" Type="Bool">true</Property>
		<Property Name="server.control.propertiesEnabled" Type="Bool">true</Property>
		<Property Name="server.tcp.enabled" Type="Bool">false</Property>
		<Property Name="server.tcp.port" Type="Int">0</Property>
		<Property Name="server.tcp.serviceName" Type="Str">我的电脑/VI服务器</Property>
		<Property Name="server.tcp.serviceName.default" Type="Str">我的电脑/VI服务器</Property>
		<Property Name="server.vi.callsEnabled" Type="Bool">true</Property>
		<Property Name="server.vi.propertiesEnabled" Type="Bool">true</Property>
		<Property Name="specify.custom.address" Type="Bool">false</Property>
		<Item Name="SetRange.vi" Type="VI" URL="../SetRange.vi"/>
		<Item Name="控件 1.ctl" Type="VI" URL="../控件 1.ctl"/>
		<Item Name="控件 4.ctl" Type="VI" URL="../控件 4.ctl"/>
		<Item Name="控件 5.ctl" Type="VI" URL="../控件 5.ctl"/>
		<Item Name="依赖关系" Type="Dependencies">
			<Item Name="vi.lib" Type="Folder">
				<Item Name="8.6CompatibleGlobalVar.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/config.llb/8.6CompatibleGlobalVar.vi"/>
				<Item Name="Application Directory.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/file.llb/Application Directory.vi"/>
				<Item Name="Check if File or Folder Exists.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/libraryn.llb/Check if File or Folder Exists.vi"/>
				<Item Name="Clear Errors.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Clear Errors.vi"/>
				<Item Name="Error Cluster From Error Code.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Error Cluster From Error Code.vi"/>
				<Item Name="NI_FileType.lvlib" Type="Library" URL="/&lt;vilib&gt;/Utility/lvfile.llb/NI_FileType.lvlib"/>
				<Item Name="NI_LVConfig.lvlib" Type="Library" URL="/&lt;vilib&gt;/Utility/config.llb/NI_LVConfig.lvlib"/>
				<Item Name="NI_PackedLibraryUtility.lvlib" Type="Library" URL="/&lt;vilib&gt;/Utility/LVLibp/NI_PackedLibraryUtility.lvlib"/>
				<Item Name="Space Constant.vi" Type="VI" URL="/&lt;vilib&gt;/dlg_ctls.llb/Space Constant.vi"/>
				<Item Name="Trim Whitespace.vi" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/Trim Whitespace.vi"/>
				<Item Name="whitespace.ctl" Type="VI" URL="/&lt;vilib&gt;/Utility/error.llb/whitespace.ctl"/>
			</Item>
		</Item>
		<Item Name="程序生成规范" Type="Build">
			<Item Name="SetRange" Type="EXE">
				<Property Name="App_copyErrors" Type="Bool">true</Property>
				<Property Name="App_INI_aliasGUID" Type="Str">{AD9C9F50-3C11-4DA6-B5D4-21A073DBBF3C}</Property>
				<Property Name="App_INI_GUID" Type="Str">{8B0D4968-419C-4CAB-A2A7-780FFD69E93C}</Property>
				<Property Name="App_serverConfig.httpPort" Type="Int">8002</Property>
				<Property Name="App_winsec.description" Type="Str">http://www.china.com</Property>
				<Property Name="Bld_autoIncrement" Type="Bool">true</Property>
				<Property Name="Bld_buildCacheID" Type="Str">{B8463ACD-C382-4858-839D-037E4A0ECE25}</Property>
				<Property Name="Bld_buildSpecName" Type="Str">SetRange</Property>
				<Property Name="Bld_defaultLanguage" Type="Str">ChineseS</Property>
				<Property Name="Bld_excludeInlineSubVIs" Type="Bool">true</Property>
				<Property Name="Bld_excludeLibraryItems" Type="Bool">true</Property>
				<Property Name="Bld_excludePolymorphicVIs" Type="Bool">true</Property>
				<Property Name="Bld_localDestDir" Type="Path">/D</Property>
				<Property Name="Bld_modifyLibraryFile" Type="Bool">true</Property>
				<Property Name="Bld_previewCacheID" Type="Str">{CB0E3A49-EDDA-433D-BB17-56C0C1EC1D7B}</Property>
				<Property Name="Bld_version.build" Type="Int">16</Property>
				<Property Name="Bld_version.major" Type="Int">1</Property>
				<Property Name="Destination[0].destName" Type="Str">SetRange.exe</Property>
				<Property Name="Destination[0].path" Type="Path">/D/SetRange.exe</Property>
				<Property Name="Destination[0].path.type" Type="Str">&lt;none&gt;</Property>
				<Property Name="Destination[0].preserveHierarchy" Type="Bool">true</Property>
				<Property Name="Destination[0].type" Type="Str">App</Property>
				<Property Name="Destination[1].destName" Type="Str">支持目录</Property>
				<Property Name="Destination[1].path" Type="Path">/D/data</Property>
				<Property Name="Destination[1].path.type" Type="Str">&lt;none&gt;</Property>
				<Property Name="DestinationCount" Type="Int">2</Property>
				<Property Name="Exe_actXinfo_enumCLSID[0]" Type="Str">{CB404689-36D1-47CC-91C2-70D0EE42E4A8}</Property>
				<Property Name="Exe_actXinfo_enumCLSID[1]" Type="Str">{225D040A-378A-48CD-896E-CFF0F39D60E5}</Property>
				<Property Name="Exe_actXinfo_enumCLSID[10]" Type="Str">{5C8275C0-C10A-4236-87F3-604AB7CA46BC}</Property>
				<Property Name="Exe_actXinfo_enumCLSID[11]" Type="Str">{48F094D4-D62E-4F21-9548-F70A1DB5C5C2}</Property>
				<Property Name="Exe_actXinfo_enumCLSID[12]" Type="Str">{057F2DD6-0C81-40BF-89ED-1B0956D05EE1}</Property>
				<Property Name="Exe_actXinfo_enumCLSID[13]" Type="Str">{541804C6-5DB7-440D-AE57-8B431F71D3A9}</Property>
				<Property Name="Exe_actXinfo_enumCLSID[14]" Type="Str">{9AAA9F23-911B-4E41-802A-3672E6B41B69}</Property>
				<Property Name="Exe_actXinfo_enumCLSID[15]" Type="Str">{BCE1E326-1B96-4735-A525-E86370C045CB}</Property>
				<Property Name="Exe_actXinfo_enumCLSID[16]" Type="Str">{DB6D7344-AE07-4A67-A0CA-9A324733B8E3}</Property>
				<Property Name="Exe_actXinfo_enumCLSID[2]" Type="Str">{6C425320-7641-440D-A2C7-3FEE9EAB5CEE}</Property>
				<Property Name="Exe_actXinfo_enumCLSID[3]" Type="Str">{B6185C09-F7F9-4868-BB63-BD61CE2A8510}</Property>
				<Property Name="Exe_actXinfo_enumCLSID[4]" Type="Str">{00CDC302-87F7-4FD7-B0ED-A324B5134D9D}</Property>
				<Property Name="Exe_actXinfo_enumCLSID[5]" Type="Str">{616E0E7A-D16C-44A3-97AE-7024C25B534B}</Property>
				<Property Name="Exe_actXinfo_enumCLSID[6]" Type="Str">{399774B3-2386-42A9-9ACB-00766C2C5B1E}</Property>
				<Property Name="Exe_actXinfo_enumCLSID[7]" Type="Str">{72DB0429-687F-4407-99C0-B4743C0086B9}</Property>
				<Property Name="Exe_actXinfo_enumCLSID[8]" Type="Str">{2DDE7D60-21DA-4764-8B3E-241C45423DE3}</Property>
				<Property Name="Exe_actXinfo_enumCLSID[9]" Type="Str">{2C0AF337-45E5-440C-9C04-DE8B51E85217}</Property>
				<Property Name="Exe_actXinfo_enumCLSIDsCount" Type="Int">17</Property>
				<Property Name="Exe_actXinfo_majorVersion" Type="Int">5</Property>
				<Property Name="Exe_actXinfo_minorVersion" Type="Int">5</Property>
				<Property Name="Exe_actXinfo_objCLSID[0]" Type="Str">{C8023ED2-AE40-458B-9252-775BFC479B68}</Property>
				<Property Name="Exe_actXinfo_objCLSID[1]" Type="Str">{C896CCFB-DBEB-49DE-A78F-53B11FE57414}</Property>
				<Property Name="Exe_actXinfo_objCLSID[10]" Type="Str">{A0FABB81-BD9F-470D-A624-602100040D68}</Property>
				<Property Name="Exe_actXinfo_objCLSID[11]" Type="Str">{82879593-2279-4A8D-B634-E5889B86F697}</Property>
				<Property Name="Exe_actXinfo_objCLSID[12]" Type="Str">{4C49AB10-1416-49DF-8FD0-FFFD949D218C}</Property>
				<Property Name="Exe_actXinfo_objCLSID[13]" Type="Str">{AB2DF2F4-52A1-4235-A804-ADEB1470B688}</Property>
				<Property Name="Exe_actXinfo_objCLSID[2]" Type="Str">{4A62BD3E-1D8C-4071-9D8A-549E7026F8F9}</Property>
				<Property Name="Exe_actXinfo_objCLSID[3]" Type="Str">{2285A156-8E3D-4F57-885D-134FD4F8993C}</Property>
				<Property Name="Exe_actXinfo_objCLSID[4]" Type="Str">{AF068696-B8A9-4290-9CB7-1E79DE18C734}</Property>
				<Property Name="Exe_actXinfo_objCLSID[5]" Type="Str">{5BFDD572-E24B-4703-9BE1-C95430409DE8}</Property>
				<Property Name="Exe_actXinfo_objCLSID[6]" Type="Str">{A1DE38CB-800F-4318-AC02-197CEA8D1078}</Property>
				<Property Name="Exe_actXinfo_objCLSID[7]" Type="Str">{E6746721-2B66-4B43-B48A-CB3000D34A38}</Property>
				<Property Name="Exe_actXinfo_objCLSID[8]" Type="Str">{A1FE8A36-429F-444D-9984-ABC12A2894CC}</Property>
				<Property Name="Exe_actXinfo_objCLSID[9]" Type="Str">{5EECB119-A5C5-4478-BF4B-38D6B0453034}</Property>
				<Property Name="Exe_actXinfo_objCLSIDsCount" Type="Int">14</Property>
				<Property Name="Exe_actXinfo_progIDPrefix" Type="Str">SetRange</Property>
				<Property Name="Exe_actXServerName" Type="Str">SetRange</Property>
				<Property Name="Exe_actXServerNameGUID" Type="Str"></Property>
				<Property Name="Source[0].itemID" Type="Str">{386444B4-14CA-48B4-BC87-244B5B16C861}</Property>
				<Property Name="Source[0].type" Type="Str">Container</Property>
				<Property Name="Source[1].destinationIndex" Type="Int">0</Property>
				<Property Name="Source[1].itemID" Type="Ref">/我的电脑/SetRange.vi</Property>
				<Property Name="Source[1].sourceInclusion" Type="Str">TopLevel</Property>
				<Property Name="Source[1].type" Type="Str">VI</Property>
				<Property Name="SourceCount" Type="Int">2</Property>
				<Property Name="TgtF_companyName" Type="Str">china</Property>
				<Property Name="TgtF_fileDescription" Type="Str">SetRange</Property>
				<Property Name="TgtF_internalName" Type="Str">SetRange</Property>
				<Property Name="TgtF_legalCopyright" Type="Str">版权 2018 china</Property>
				<Property Name="TgtF_productName" Type="Str">SetRange</Property>
				<Property Name="TgtF_targetfileGUID" Type="Str">{5510DA2A-DAF9-4785-AAFD-9B7AAF23DC14}</Property>
				<Property Name="TgtF_targetfileName" Type="Str">SetRange.exe</Property>
			</Item>
		</Item>
	</Item>
</Project>
