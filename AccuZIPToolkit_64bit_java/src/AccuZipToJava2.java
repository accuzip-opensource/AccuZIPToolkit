import java.io.File;

import com.servoyguy.plugins.servoycom.JSVariant;
import com.servoyguy.plugins.servoycom.RemoteCOM;

/**
 *
 * reference => http://www.accuzip.com/files/AccuZIPToolkit.pdf
 */
public class AccuZipToJava2 {

	private static boolean loadedJacobDll = false;

	/**
	 * initializer
	 */
	static {
		/**
		 * AbsolutePath of jacob-1.18-x64.dll
		 */

		File jacobDLL = new File("jacob-1.18-x64.dll");
		loadedJacobDll = jacobDLL.exists();
		String filePath = jacobDLL.getAbsolutePath();
		System.load(filePath);
	}

	private RemoteCOM rc = new RemoteCOM("AccuAddress.AddrCode6");
	private File my_config_acu_file = null;

	
	/**
	 * Constructor
	 * @param path_to_Config_acu
	 * @throws Exception
	 */
	public AccuZipToJava2(String path_to_Config_acu) throws Exception {

		this.my_config_acu_file = new File(path_to_Config_acu);
		if (!this.my_config_acu_file.exists()) {
			throw new Exception(
					"Cannot find Config.acu in the path you provided => "
							+ this.my_config_acu_file.getAbsolutePath());
		}
		jacobDllLoadStatus();
		openAccuAddress();
	}

	/**
	 * Reference: http://www.accuzip.com/files/AccuZIPToolkit.pdf see â€œGETâ€�
	 * Function Calls for environments that do not support â€œbyRefâ€�
	 * 
	 * @throws Exception
	 * 
	 */
	public void processAddress() throws Exception {
		
		/**
		 * Address Inputs
		 * 
		 * Set Company Name AddrCode.AZSetQuery_iadl2("Company Name")
		 * 
		 * Set Address Line 1 AddrCode.AZSetQuery_iadl1("Primary Address")
		 * 
		 * Set Address Line 2
		 * AddrCode.AZSetQuery_iadl3("Secondary Address or Suite/Apt")
		 * 
		 * Set city Name AddrCode.AZSetQuery_ictyi("City")
		 * 
		 * Set state AddrCode.AZSetQuery_istai("State")
		 * 
		 * Set ZIP Code AddrCode.AZSetQuery_izipc("ZIP or ZIP+4")
		 * 
		 * Set Urbanization AddrCode.AZSetQuery_iprurb("Urbanization")
		 */

		this.rc.call("AZSetQuery_iadl2", "ACCUZIP");
		this.rc.call("AZSetQuery_iadl1", "3216 El Camino Real");
		this.rc.call("AZSetQuery_iadl3", "");
		this.rc.call("AZSetQuery_ictyi", "Atascadero");
		this.rc.call("AZSetQuery_istai", "CA");
		this.rc.call("AZSetQuery_izipc", "93422");
		this.rc.call("AZSetQuery_iprurb", "");

		this.rc.call("AZFindDeliverable");

		/**
		 * Address responses
		 * 
		 * Standardized Primary Address dadl1=AddrCode.AZGetResult_dadl1_AX()
		 * 
		 * Standardized Input City (if acceptable)
		 * dctya=AddrCode.AZGetResult_dctya_AX()
		 * 
		 * Standardized State Abbreviation dstaa=AddrCode.AZGetResult_dstaa_AX()
		 * 
		 * Standardized ZIP Code zipc=AddrCode.AZGetResult_zipc_AX()
		 * 
		 * +4 Code addon=AddrCode.AZGetResult_addon_AX()
		 * 
		 * Parsed Pre-direction pre_dir=AddrCode.AZGetAddr_pre_dir_AX()
		 * 
		 * Parsed Street Name str_name=AddrCode.AZGetAddr_str_name_AX()
		 * 
		 * Parsed Suffix suffix=AddrCode.AZGetAddr_suffix_AX()
		 * 
		 * Parsed Post Direction post_dir=AddrCode.AZGetAddr_post_dir_AX()
		 * 
		 * Parsed Unit Description unit=AddrCode.AZGetAddr_unit_AX()
		 * 
		 * Parsed Secondary Number msnum=AddrCode.AZGetResult_msnum_AX()
		 * 
		 * Private Mail Box Description pmbunit=AddrCode.AZGetResult_pmb_AX()
		 * 
		 * Private Mail Box Number pmbnum=AddrCode.AZGetResult_pmbnum_AX()
		 */

		/**
		 * response in JSON format
		 */
		JSVariant jsonStruct = this.rc.call("AZJSONFormatResult");
		String jsonString = jsonStruct.toString();
		System.out.println("response jsonString => " + jsonString);

		JSVariant dadl1Result = this.rc.call("AZGetResult_dadl1_AX");
		JSVariant dctyaResult = this.rc.call("AZGetResult_dctya_AX");
		JSVariant dstaaResult = this.rc.call("AZGetResult_dstaa_AX");
		JSVariant zipcResult = this.rc.call("AZGetResult_zipc_AX");
		JSVariant addonResult = this.rc.call("AZGetResult_addon_AX");

		JSVariant pre_dirResult = this.rc.call("AZGetAddr_pre_dir_AX");
		JSVariant strnameResult = this.rc.call("AZGetAddr_str_name_AX");
		JSVariant suffixResult = this.rc.call("AZGetAddr_suffix_AX");

		JSVariant post_dirResult = this.rc.call("AZGetAddr_post_dir_AX");
		JSVariant unitResult = this.rc.call("AZGetAddr_unit_AX");
		JSVariant msnumResult = this.rc.call("AZGetResult_msnum_AX");

		String errorStr = this.rc.getLastError();
		if (errorStr != null && errorStr.trim().length() != 0) {
			System.out.println("errorStr => " + errorStr);
		}

		System.out.println("dadl1Result => " + dadl1Result);
		System.out.println("dctyaResult => " + dctyaResult);
		System.out.println("dstaaResult => " + dstaaResult);
		System.out.println("zipcResult => " + zipcResult);
		System.out.println("addonResult => " + addonResult);

		System.out.println("pre_dirResult => " + pre_dirResult);
		System.out.println("strnameResult => " + strnameResult);
		System.out.println("suffixResult => " + suffixResult);
		System.out.println("post_dirResult => " + post_dirResult);
		System.out.println("unitResult => " + unitResult);
		System.out.println("msnumResult => " + msnumResult);

		dadl1Result = null;
		dctyaResult = null;
		dstaaResult = null;
		zipcResult = null;
		addonResult = null;

		pre_dirResult = null;
		strnameResult = null;
		suffixResult = null;

		post_dirResult = null;
		unitResult = null;
		msnumResult = null;
		
	}

	/**
	 * Number of days until USPS Data Expires (Optional call)
	 * nResult=AddrCode.zp4GetDataExpirationDays("PathTo\Config.acu",0)
	 * @throws Exception 
	 */
	public void zp4GetDataExpirationDays() throws Exception {
		
		JSVariant nResult2 = rc.call("zp4GetDataExpirationDays", this.my_config_acu_file.getAbsolutePath(), 0);
		if (nResult2 == null) {
			String errorStr = this.rc.getLastError();
			throw new Exception("RemoteCOM zp4GetDataExpirationDays nResult = " + nResult2	+ "\n Error: => " + errorStr);
		}
				

		/**
		 * If 14 or fewer days remain,  Alert.
		 */
		
		if (nResult2.getInt() < 15) {
			System.out.println("expDays < 15: days left => " + nResult2.getInt());
		}
		System.out.println("expDays => " + nResult2.getInt());
	}

	private void jacobDllLoadStatus() throws Exception {
		if (!loadedJacobDll) {
			throw new Exception("Problem Loading Jacob dll");
		}
	}

	public void openAccuAddress() throws Exception {

		/**
		 * nResult=AddrCode.zp4connectalt("PathTo\Config.acu")
		 * 
		 * return values for nResult
		 *  0 No errors, opened successfully 
		 *  1 Error opening a file 
		 *  2 Error reading a file 
		 *  3 Error writing to Plus4cfg.dat file located in NATIONAL folder
		 *  4 Error finding a file
		 *  5 Registration Activation Number or Database has expired 6 Database files are
		 */

		JSVariant nResult = this.rc.call("zp4connectalt", this.my_config_acu_file.getAbsolutePath());

		if (nResult == null || nResult.getInt() != 0) {
			String errorStr = this.rc.getLastError();
			throw new Exception("RemoteCOM zp4connectalt nResult = " + nResult
					+ "\n Error: => " + errorStr);
		}
	}

	public void closeAccuAddress() {

		this.rc.call("zp4close");
	}

	public static void main(String[] args) throws Exception {

		AccuZipToJava2 az2 = null;

		try {
			/**
			 * enter the path to Config.acu => new AccuZipToJava2("path to Config.acu");
			 * 
			 */
			az2 = new AccuZipToJava2("c:\\Program Files (x86)\\AccuZIP6 5.0\\Config.acu");

			az2.openAccuAddress();

			az2.processAddress();

		} catch (Exception e) {
			throw e;
		} finally {
			/**
			 * always call rc.call("zp4close"); when finished.
			 */
			if (az2 != null) {
				az2.closeAccuAddress();
			}
		}
	}

}
