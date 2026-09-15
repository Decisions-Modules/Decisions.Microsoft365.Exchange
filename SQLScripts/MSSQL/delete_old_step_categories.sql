/* ---------------------------------------------------------------- */
/*            Script to Delete Steps with Old Categories            */
/* ---------------------------------------------------------------- */
/* Notes:                                                           */
/* Run the Delete statements in a transaction                       */


IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'element_registration' and TABLE_SCHEMA = SCHEMA_NAME())
BEGIN
BEGIN TRY

DELETE dbo.element_registration
FROM dbo.element_registration
WHERE element_registration.class_name In ('Decisions.Microsoft365.Exchange.Steps.CalendarSteps','Decisions.Microsoft365.Exchange.Steps.ContactSteps','Decisions.Microsoft365.Exchange.Steps.EmailSteps','Decisions.Microsoft365.Exchange.Steps.GroupSteps')

END TRY
BEGIN CATCH
-- Best-effort cleanup; do not let a transient failure here (e.g. a
-- dependency not yet created during first-boot module install) abort
-- the caller's ambient transaction. Do not BEGIN/COMMIT/ROLLBACK here:
-- this script runs inside a transaction owned by the caller, and an
-- unqualified ROLLBACK would unwind that transaction too.
END CATCH
END