/**
 * README
 * This extension is used to delete staggered prices in the MPAGRP table
 * Name: EXT018MI.DltStgPrice
 * Description: Delete staggered price records in the MPAGRP table
 * Date       Changed By           Description
 * 20260312   EPRONOVOST           Added DltStgPrice transaction to delete purchase agreement staggered prices
 * 20260326   EPRONOVOST           Fixed header description; changed readAllLock to readAll + readLock pattern
 * 20260407   EPRONOVOST           Limit nrOfRecords to 10000
 */
import java.time.LocalDate
import java.time.format.DateTimeFormatter
import java.time.format.DateTimeParseException

public class DltStgPrice extends ExtendM3Transaction {
  private final MIAPI mi
  private final DatabaseAPI database
  private final LoggerAPI logger
  private final ProgramAPI program

  // Global variables to store input parameters for use across closures
  private String iCONO
  private String iSUNO
  private String iAGNB
  private String iGRPI
  private String iOBV1
  private String iOBV2
  private String iFVDT
  private String iFRQT

  public DltStgPrice(MIAPI mi, DatabaseAPI database, LoggerAPI logger, ProgramAPI program) {
    this.mi = mi
    this.database = database
    this.logger = logger
    this.program = program
  }

  /**
   * Main entry point for the transaction.
   * Handles input extraction, null-safe defaults, and rigorous date validation.
   */
  public void main() {
    // Determine Company Number (CONO); defaults to current user company if blank
    iCONO = mi.inData.get("CONO") == null || mi.inData.get("CONO").isBlank() ? program.getLDAZD().CONO as String : mi.inData.get("CONO")
    
    // Retrieve inputs and apply trimming to prevent lookup failures due to whitespace
    iSUNO = mi.inData.get("SUNO") == null || mi.inData.get("SUNO").isBlank() ? "" : mi.inData.get("SUNO").trim()
    iAGNB = mi.inData.get("AGNB") == null || mi.inData.get("AGNB").isBlank() ? "" : mi.inData.get("AGNB").trim()
    iGRPI = mi.inData.get("GRPI") == null || mi.inData.get("GRPI").isBlank() ? "" : mi.inData.get("GRPI").trim()
    iOBV1 = mi.inData.get("OBV1") == null || mi.inData.get("OBV1").isBlank() ? "" : mi.inData.get("OBV1").trim()
    iOBV2 = mi.inData.get("OBV2") == null || mi.inData.get("OBV2").isBlank() ? "" : mi.inData.get("OBV2").trim()
    iFVDT = mi.inData.get("FVDT") == null || mi.inData.get("FVDT").isBlank() ? "" : mi.inData.get("FVDT").trim()
    iFRQT = mi.inData.get("FRQT") == null || mi.inData.get("FRQT").isBlank() ? "" : mi.inData.get("FRQT").trim()

    // Mandatory Field Validation
    if (iSUNO.isBlank()) {
      mi.error("Supplier number (SUNO) is required")
      return
    }

    if (iFVDT.isBlank()) {
      mi.error("From valid date (FVDT) is required")
      return
    }

    // Ensure the date string is a valid M3 date format before attempting DB operations
    if (!isDateValid(iFVDT, "yyyyMMdd")) {
      mi.error("From valid date (FVDT) must be in yyyyMMdd format")
      return
    }

    delStgPrice()
  }

  /**
   * Identifies records in the MPAGRP table (Purchase Agreement Staggered Price).
   * Utilizes 10 keys (Index 00) to isolate specific quantity-based price breaks.
   */
  private void delStgPrice() {
    DBAction action = database.table("MPAGRP").index("00").build()
    DBContainer container = action.getContainer()
    
    // Set all 10 key fields for the primary index
    container.set("AJCONO", iCONO.toInteger())
    container.set("AJSUNO", iSUNO)
    container.set("AJAGNB", iAGNB)
    container.set("AJGRPI", iGRPI.toInteger())
    container.set("AJOBV1", iOBV1)
    container.set("AJOBV2", iOBV2)
    container.set("AJOBV3", "") // Empty object values required for full key match
    container.set("AJOBV4", "")
    container.set("AJFVDT", iFVDT.toInteger())
    container.set("AJFRQT", iFRQT.toDouble()) // 'From Quantity' is the final differentiator
    
    // Search using a full 10-key match
    action.readAll(container, 10, 10000, readStgPriceCallBack)
  }

  /**
   * Callback for MPAGRP search. 
   * Prepares the record for deletion by establishing a row-level lock.
   */
  Closure<?> readStgPriceCallBack = { DBContainer row ->
    DBAction lockAction = database.table("MPAGRP").index("00").build()
    DBContainer lockContainer = lockAction.getContainer()
    
    // Map existing row data to a new container for the Lock request
    lockContainer.set("AJCONO", row.getInt("AJCONO"))
    lockContainer.set("AJSUNO", row.getString("AJSUNO"))
    lockContainer.set("AJAGNB", row.getString("AJAGNB"))
    lockContainer.set("AJGRPI", row.getInt("AJGRPI"))
    lockContainer.set("AJOBV1", row.getString("AJOBV1"))
    lockContainer.set("AJOBV2", row.getString("AJOBV2"))
    lockContainer.set("AJOBV3", row.getString("AJOBV3"))
    lockContainer.set("AJOBV4", row.getString("AJOBV4"))
    lockContainer.set("AJFVDT", row.getInt("AJFVDT"))
    lockContainer.set("AJFRQT", row.getDouble("AJFRQT"))
    
    // Request exclusive lock to perform the delete operation
    lockAction.readLock(lockContainer, deleteStgPriceCallBack)
  }

  /**
   * Final step: Deletes the locked record from MPAGRP.
   */
  Closure<?> deleteStgPriceCallBack = { LockedResult lockedResult ->
    lockedResult.delete()
  }

  /**
   * Utility method to validate date format using Java 8 Time API.
   * @param date The string representation of the date.
   * @param format The expected pattern (e.g., yyyyMMdd).
   * @return true if valid, false otherwise.
   */
  private boolean isDateValid(String date, String format) {
    try {
      LocalDate.parse(date, DateTimeFormatter.ofPattern(format))
      return true
    } catch (DateTimeParseException e) {
      // Log error if necessary via logger.debug/error
      return false
    }
  }
}