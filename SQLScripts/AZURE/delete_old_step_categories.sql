/* ---------------------------------------------------------------- */
/*            Script to Delete Steps with Old Categories            */
/* ---------------------------------------------------------------- */
/* Notes:                                                           */
/* Run the Delete statements in a transaction                       */

IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'element_registration' and TABLE_SCHEMA = SCHEMA_NAME())
BEGIN
BEGIN TRANSACTION

DELETE dbo.element_registration
FROM dbo.element_registration
WHERE element_registration.class_name In ('Decisions.Microsoft365.Exchange.Steps.CalendarSteps','Decisions.Microsoft365.Exchange.Steps.ContactSteps','Decisions.Microsoft365.Exchange.Steps.EmailSteps','Decisions.Microsoft365.Exchange.Steps.GroupSteps')


COMMIT
END