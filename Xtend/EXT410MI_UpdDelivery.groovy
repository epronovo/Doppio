/**
 * README
 * This extension is being used to update details in the MHDISH table
 * Name: EXT410MI.UpdDelivery
 * Description: Update details in MHDISH table
 * Date      Changed By            Description
 * 20260728  EPRONOVOST            Added UpdDelivery transaction to access MHDISH table
 */

import java.time.LocalDate
import java.time.format.DateTimeFormatter

public class UpdDelivery extends ExtendM3Transaction {
	private final MIAPI mi
	private final DatabaseAPI database
	private final ProgramAPI program

	private String iCONO
	private String iINOU
	private String iDLIX
	private String iRLTD

	public UpdDelivery(MIAPI mi, DatabaseAPI database, ProgramAPI program	) {
		this.mi = mi
		this.database = database
		this.program = program
	}

	/**
	 * Entry point. Reads CONO, INOU, DLIX, and RLTD from the input parameters
	 * and updates the corresponding record in MHDISH.
	 */
	public void main() {
		String rawCONO = mi.inData.get("CONO") ?: ""
		String rawINOU = mi.inData.get("INOU") ?: ""
		String rawDLIX = mi.inData.get("DLIX") ?: ""
		String rawRLTD = mi.inData.get("RLTD") ?: ""

		iCONO = rawCONO.isBlank() ? program.getLDAZD().CONO : rawCONO
		iINOU = rawINOU.isBlank() ? "" : rawINOU
		iDLIX = rawDLIX.isBlank() ? "" : rawDLIX
		iRLTD = rawRLTD.isBlank() ? "" : rawRLTD

		if (!validateInput()) {
			return
		}

		updRecord()
	}

	/**
	 * Validates the numeric business rules for CONO, INOU, DLIX, and RLTD.
	 * Reports an error and returns false on the first rule violated.
	 */
	private boolean validateInput() {
		if (iCONO.toInteger() == 0) {
			mi.error("CONO must not be 0")
			return false
		}

		int inou = iINOU.toInteger()
		if (inou < 1 || inou > 4) {
			mi.error("INOU must be between 1 and 4")
			return false
		}

		if (iDLIX.toInteger() == 0) {
			mi.error("DLIX must not be 0")
			return false
		}

		if (!iRLTD.isBlank() && iRLTD.trim() != "?") {
			int rltd = iRLTD.toInteger()
			if (rltd != 0 && rltd != 1) {
				mi.error("RLTD must be 0 or 1")
				return false
			}
		}

		return true
	}

	/**
	 * Looks up the MHDISH record by CONO/INOU/DLIX and locks it for update
	 * via updateCallBack. Reports an error if no matching record exists.
	 */
	private void updRecord() {
		DBAction action = database.table("MHDISH").index("00").build()
		DBContainer container = action.getContainer()

		container.set("OQCONO", iCONO.toInteger())
		container.set("OQINOU", iINOU.toInteger())
		container.set("OQDLIX", iDLIX.toInteger())

		if (!action.readLock(container, updateCallBack)) {
			mi.error("Record does not exist")
		}
	}

	/**
	 * Applies the update to the locked MHDISH record: sets RLTD (if provided),
	 * refreshes OQLMDT/OQCHID, and increments OQCHNO (wrapping to 1 after 999).
	 */
	Closure <?> updateCallBack = {
		LockedResult lockedResult ->

		int changeNumber = lockedResult.get("OQCHNO")
	    int newChangeNumber = changeNumber >= 999 ? 1 : changeNumber + 1

		if (!iRLTD.isBlank()) {
			if (iRLTD.trim() == "?") {iRLTD = "0"}
			lockedResult.set("OQRLTD", iRLTD.toInteger())
		}

		lockedResult.set("OQLMDT", LocalDate.now().format(DateTimeFormatter.ofPattern("yyyyMMdd")).toInteger())
		lockedResult.set("OQCHNO", newChangeNumber)
		lockedResult.set("OQCHID", program.getUser())

		lockedResult.update()
	}
}
