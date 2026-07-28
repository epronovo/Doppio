/**
 * README
 * This extension is being used to update details in the MITMAS table
 * Name: EXT014MI.UpdItem
 * Description: Update details in MITMAS table
 * Date      Changed By            Description
 * 20241125  EPRONOVOST            Added UpdItem transaction to access MITMAS table
 * 20250130  EPRONOVOST            Changes to mandatory input fields, added expired date
 * 20250417  EPRONOVOST            Changed expiration date to include May
 * 20250812  EPRONOVOST            Changed expiration date to include October
 * 20260402  EPRONOVOST            Changed expiration date to include Mid June
 */

import java.time.LocalDate
import java.time.format.DateTimeFormatter

public class UpdItem extends ExtendM3Transaction {
	private final MIAPI mi
	private final DatabaseAPI database
	private final LoggerAPI logger
	private final ProgramAPI program
	private final UtilityAPI utility

	private String iCONO
	private String iITNO
	private String iBACD

	public UpdItem(MIAPI mi, DatabaseAPI database, LoggerAPI logger, ProgramAPI program, UtilityAPI utility) {
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
	  
    //Expiration Date for data correction extension
    if (LocalDate.now().isAfter(LocalDate.of(2026, 6, 15))) {
    	mi.error("Extension signature expired ${program.LDAZD.TIZO}")
    	logger.debug("Extension signature expired")
    	return
    }

		iCONO = mi.inData.get("CONO").isBlank() ? program.getLDAZD().CONO : mi.inData.get("CONO")
		iITNO = mi.inData.get("ITNO").isBlank() ? "" : mi.inData.get("ITNO")
		iBACD = mi.inData.get("BACD").isBlank() ? "" : mi.inData.get("BACD")

		updRecord()
	}

	/**
	 * Update the record in MITMAS table
	 *
	 */
	private void updRecord() {
		DBAction action = database.table("MITMAS").index("00").build()
		DBContainer container = action.getContainer()

		container.set("MMCONO", iCONO.toInteger())
		container.set("MMITNO", iITNO)

		if (!action.readLock(container, updateCallBack)) {
			mi.error("Record does not exist")
			return
		}
	}

	/**
	 * Closure for MITMAS Update
	 *
	 */
	Closure <?> updateCallBack = {
		LockedResult lockedResult ->

		int changeNumber = lockedResult.get("MMCHNO")
		int newChangeNumber = changeNumber + 1

		if (!iBACD.isBlank()) {
			if (iBACD.trim() == "?") {iBACD = "0"}
			lockedResult.set("MMBACD", iBACD.toInteger())
		}

		lockedResult.set("MMLMDT", LocalDate.now().format(DateTimeFormatter.ofPattern("yyyyMMdd")).toInteger())
		lockedResult.set("MMCHNO", newChangeNumber)
		lockedResult.set("MMCHID", program.getUser())

		lockedResult.update()
	}
}
