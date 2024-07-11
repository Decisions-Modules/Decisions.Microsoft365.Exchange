/* ---------------------------------------------------------------- */
/*            Script to Delete Steps with Old Categories            */
/* ---------------------------------------------------------------- */
/* Notes:                                                           */
/* Run the Delete statements in a transaction                       */


DO $$ 
BEGIN
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_name = 'element_registration' AND table_schema = current_schema()) THEN
        DELETE FROM element_registration
        WHERE class_name IN ('Decisions.Microsoft365.Exchange.Steps.CalendarSteps',
                             'Decisions.Microsoft365.Exchange.Steps.ContactSteps',
                             'Decisions.Microsoft365.Exchange.Steps.EmailSteps',
                             'Decisions.Microsoft365.Exchange.Steps.GroupSteps');
    END IF;
END $$;