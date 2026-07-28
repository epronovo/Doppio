/**
 * README
 * This extension is used to delete agreement lines and their linked staggered prices
 * Name: EXT018MI.DltAgrLine
 * Description: Delete records in MPAGRL table and all linked MPAGRP records
 * Date       Changed By           Description
 * 20260312   EPRONOVOST           Added DltAgrLine transaction to delete purchase agreement lines and linked staggered prices
 * 20260326   EPRONOVOST           Fixed header description; changed readAllLock to readAll + readLock pattern
 * 20260407   EPRONOVOST           Limit nrOfRecords to 10000
 */
public class DltAgrLine extends ExtendM3Transaction {
  private final MIAPI mi
  private final DatabaseAPI database
  private final LoggerAPI logger
  private final ProgramAPI program

  // Input fields from MI transaction
  private String iCONO
  private String iSUNO
  private String iAGNB
  private String iGRPI
  private String iOBV1
  private String iOBV2
  private String iFVDT

  public DltAgrLine(MIAPI mi, DatabaseAPI database, LoggerAPI logger, ProgramAPI program) {
    this.mi = mi
    this.database = database
    this.logger = logger
    this.program = program
  }

  /**
   * Main entry point for the MI transaction.
   * Handles input retrieval, default values, and mandatory field validation.
   */
  public void main() {
    // Get Company Number (CONO); default to user's current company if not provided
    iCONO = mi.inData.get("CONO") == null || mi.inData.get("CONO").isBlank() ? program.getLDAZD().CONO as String : mi.inData.get("CONO")
    
    // Retrieve and trim input parameters
    iSUNO = mi.inData.get("SUNO") == null || mi.inData.get("SUNO").isBlank() ? "" : mi.inData.get("SUNO").trim()
    iAGNB = mi.inData.get("AGNB") == null || mi.inData.get("AGNB").isBlank() ? "" : mi.inData.get("AGNB").trim()
    iGRPI = mi.inData.get("GRPI") == null || mi.inData.get("GRPI").isBlank() ? "" : mi.inData.get("GRPI").trim()
    iOBV1 = mi.inData.get("OBV1") == null || mi.inData.get("OBV1").isBlank() ? "" : mi.inData.get("OBV1").trim()
    iOBV2 = mi.inData.get("OBV2") == null || mi.inData.get("OBV2").isBlank() ? "" : mi.inData.get("OBV2").trim()
    iFVDT = mi.inData.get("FVDT") == null || mi.inData.get("FVDT").isBlank() ? "" : mi.inData.get("FVDT").trim()

    // Mandatory field check: Supplier Number is critical for identifying the agreement
    if (iSUNO.isBlank()) {
      mi.error("Supplier number (SUNO) is required")
      return
    }

    delAgrLine()
  }

  /**
   * Initiates the deletion process by querying MPAGRL (Purchase Agreement Lines).
   * Queries the table using 9 keys to find specific lines to remove.
   */
  private void delAgrLine() {
    DBAction action = database.table("MPAGRL").index("00").build()
    DBContainer container = action.getContainer()
    
    // Set search keys for MPAGRL Index 00
    container.set("AICONO", iCONO.toInteger())
    container.set("AISUNO", iSUNO)
    container.set("AIAGNB", iAGNB)
    container.set("AIGRPI", iGRPI.toInteger())
    container.set("AIOBV1", iOBV1)
    container.set("AIOBV2", iOBV2)
    container.set("AIOBV3", "") // Defaulting empty as they are part of the index
    container.set("AIOBV4", "")
    container.set("AIFVDT", iFVDT.toInteger())
    
    // Execute readAll; results are processed in readAgrLineCallBack
    action.readAll(container, 9, readAgrLineCallBack)
  }

  /**
   * Callback for MPAGRL search results.
   * Cascades the deletion: first removes child records (MPAGRP), then locks the parent line for deletion.
   */
  Closure<?> readAgrLineCallBack = { DBContainer row ->
    // Cascade Delete: Must remove staggered prices (MPAGRP) before the parent line (MPAGRL)
    delStgPrice(row)

    // Re-access the record with a Lock to ensure data integrity during deletion
    DBAction lockAction = database.table("MPAGRL").index("00").build()
    DBContainer lockContainer = lockAction.getContainer()
    lockContainer.set("AICONO", row.getInt("AICONO"))
    lockContainer.set("AISUNO", row.getString("AISUNO"))
    lockContainer.set("AIAGNB", row.getString("AIAGNB"))
    lockContainer.set("AIGRPI", row.getInt("AIGRPI"))
    lockContainer.set("AIOBV1", row.getString("AIOBV1"))
    lockContainer.set("AIOBV2", row.getString("AIOBV2"))
    lockContainer.set("AIOBV3", row.getString("AIOBV3"))
    lockContainer.set("AIOBV4", row.getString("AIOBV4"))
    lockContainer.set("AIFVDT", row.getInt("AIFVDT"))
    
    lockAction.readLock(lockContainer, deleteAgrLineCallBack)
  }

  /**
   * Final step for MPAGRL deletion.
   * Executed once a row-level lock is successfully acquired.
   */
  Closure<?> deleteAgrLineCallBack = { LockedResult lockedResult ->
    lockedResult.delete()
  }

  /**
   * Queries MPAGRP (Purchase Agreement Staggered Prices) linked to a specific agreement line.
   * @param agrlRow The parent MPAGRL container used to provide relational keys.
   */
  private void delStgPrice(DBContainer agrlRow) {
    DBAction action = database.table("MPAGRP").index("00").build()
    DBContainer container = action.getContainer()
    
    // Map MPAGRL (AI) fields to MPAGRP (AJ) fields to find all price breaks
    container.set("AJCONO", agrlRow.getInt("AICONO"))
    container.set("AJSUNO", agrlRow.getString("AISUNO"))
    container.set("AJAGNB", agrlRow.getString("AIAGNB"))
    container.set("AJGRPI", agrlRow.getInt("AIGRPI"))
    container.set("AJOBV1", agrlRow.getString("AIOBV1"))
    container.set("AJOBV2", agrlRow.getString("AIOBV2"))
    container.set("AJOBV3", agrlRow.getString("AIOBV3"))
    container.set("AJOBV4", agrlRow.getString("AIOBV4"))
    container.set("AJFVDT", agrlRow.getInt("AIFVDT"))
    
    // Read all price breaks for the current agreement line
    action.readAll(container, 9, 10000, readStgPriceCallBack)
  }

  /**
   * Callback for MPAGRP search results.
   * Acquires a lock on each individual staggered price record.
   */
  Closure<?> readStgPriceCallBack = { DBContainer row ->
    DBAction lockAction = database.table("MPAGRP").index("00").build()
    DBContainer lockContainer = lockAction.getContainer()
    lockContainer.set("AJCONO", row.getInt("AJCONO"))
    lockContainer.set("AJSUNO", row.getString("AJSUNO"))
    lockContainer.set("AJAGNB", row.getString("AJAGNB"))
    lockContainer.set("AJGRPI", row.getInt("AJGRPI"))
    lockContainer.set("AJOBV1", row.getString("AJOBV1"))
    lockContainer.set("AJOBV2", row.getString("AJOBV2"))
    lockContainer.set("AJOBV3", row.getString("AJOBV3"))
    lockContainer.set("AJOBV4", row.getString("AJOBV4"))
    lockContainer.set("AJFVDT", row.getInt("AJFVDT"))
    lockContainer.set("AJFRQT", row.getDouble("AJFRQT")) // Necessary key for the staggered price break
    
    lockAction.readLock(lockContainer, deleteStgPriceCallBack)
  }

  /**
   * Final step for MPAGRP deletion.
   * Executed once the staggered price row is locked.
   */
  Closure<?> deleteStgPriceCallBack = { LockedResult lockedResult ->
    lockedResult.delete()
  }
}