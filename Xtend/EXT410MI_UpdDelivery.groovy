/**
 * README
 * This extension is being used to update details in the MHDISH table
 * Name: EXT410MI.UpdDelivery
 * Description: Update details in MHDISH table
 * Date      Changed By            Description
 * 20260728  EPRONOVOST            Added UpdDelivery transaction to access MHDISH table
 */

// import java.time.Instant
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
	 * Main method
	 *
	 * @param
	 * @return
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
			if (iRLTD.trim() == "?") {iRLTD = "0"}
			lockedResult.set("OQRLTD", iRLTD.toInteger())
		}

		lockedResult.set("OQLMDT", LocalDate.now().format(DateTimeFormatter.ofPattern("yyyyMMdd")).toInteger())
		lockedResult.set("OQCHNO", newChangeNumber)
		lockedResult.set("OQCHID", program.getUser())

		lockedResult.update()
	}
}
