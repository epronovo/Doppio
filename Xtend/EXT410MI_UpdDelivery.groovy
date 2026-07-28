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
	private final LoggerAPI logger
	private final ProgramAPI program
	private final UtilityAPI utility

	private String iCONO
	private String iINOU
	private String iDLIX
	private String iRLTD

	public UpdDelivery(MIAPI mi, DatabaseAPI database, LoggerAPI logger, ProgramAPI program, UtilityAPI utility) {
		this.mi = mi
		this.database = database
		this.logger = logger
		this.program = program
		this.utility = utility
	}

	/**
	 * Main method
	 *
	 * @param
	 * @return
	 */
	public void main() {
		iCONO = mi.inData.get("CONO").isBlank() ? program.getLDAZD().CONO.toString() : mi.inData.get("CONO")
		iINOU = mi.inData.get("INOU").isBlank() ? "" : mi.inData.get("INOU")
		iDLIX = mi.inData.get("DLIX").isBlank() ? "" : mi.inData.get("DLIX")
		iRLTD = mi.inData.get("RLTD").isBlank() ? "" : mi.inData.get("RLTD")

		if (iINOU.isBlank()) {
			mi.error("INOU must be specified")
			return
		}

		if (iDLIX.isBlank()) {
			mi.error("DLIX must be specified")
			return
		}

		updRecord()
	}

	/**
	 * Update the record in MHDISH table
	 *
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
	 * Closure for MHDISH Update
	 *
	 */
	Closure <?> updateCallBack = {
		LockedResult lockedResult ->

		int changeNumber = lockedResult.get("OQCHNO")
		int newChangeNumber = changeNumber + 1

		if (!iRLTD.isBlank()) {
			// M3 MI API convention: "?" means clear the field, so map it to 0
			if (iRLTD.trim() == "?") {iRLTD = "0"}
			lockedResult.set("OQRLTD", iRLTD.toInteger())
		}

		// OQLMDT is stored as an int (yyyyMMdd); confirm against the MHDISH data dictionary if this ever fails
		lockedResult.set("OQLMDT", LocalDate.now().format(DateTimeFormatter.ofPattern("yyyyMMdd")).toInteger())
		lockedResult.set("OQCHNO", newChangeNumber)
		lockedResult.set("OQCHID", program.getUser())

		lockedResult.update()
	}
}
