/* ---------------------------------------------------------------- */
/*            Script to Delete Steps with Old Categories            */
/* ---------------------------------------------------------------- */
/* Notes:                                                           */
/* Run the Delete statements in a transaction                       */

-- element_configuration_data is checked because element_registration's
-- delete trigger depends on it; on a from-scratch container install this
-- module's SQL can run before that table exists yet, which dooms the
-- caller's ambient transaction (no amount of TRY/CATCH here can save a
-- doomed transaction, so skip the cleanup instead until it's safe to run).
IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'element_registration' and TABLE_SCHEMA = SCHEMA_NAME())
AND EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'element_configuration_data' and TABLE_SCHEMA = SCHEMA_NAME())
BEGIN

DELETE dbo.element_registration
FROM dbo.element_registration
WHERE element_registration.class_name In ('Decisions.Microsoft365.Exchange.Steps.CalendarSteps','Decisions.Microsoft365.Exchange.Steps.ContactSteps','Decisions.Microsoft365.Exchange.Steps.EmailSteps','Decisions.Microsoft365.Exchange.Steps.GroupSteps')

END