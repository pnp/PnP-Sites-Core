create role [PnPLimitedReader]

grant select on dbo.TestRunSet to [PnPLimitedReader]
grant select on dbo.TestResultSet to [PnPLimitedReader]
grant select on dbo.TestResultMessageSet to [PnPLimitedReader]
grant select on dbo.TestConfigurationSet to [PnPLimitedReader]
grant select on dbo.FileTrackingSet to [PnPLimitedReader]
grant select on dbo.FileTrackingBaselineSet to [PnPLimitedReader]