using System.Data;
using System.Data.SqlClient;
using System.Diagnostics.CodeAnalysis;
using DailyExtractionReports.Models;
using Dapper;
using Microsoft.Extensions.Configuration;

namespace DailyExtractionReports;

public static class DatabaseHandler
{
	public static IEnumerable<Duplicates> GetPossibleDuplicates()
	{
		var appSettings = new Startup();
		var conString = appSettings.Config?.GetSection("ConnectionStrings").GetValue<string>("windows-auth-PROD");

		using var conLabOps = new SqlConnection(conString);
		if (conLabOps.State == ConnectionState.Closed) conLabOps.Open();

		const string duplicatesQuery =
			@"--Determine if any patient specimens were extracted multiple times after a specific datetime.
		--Set the @batchCreatedDateTimeStart to what is desired and it will query from that day forward till present.
		--Ideally in the future an end date could be added to allow historic sections of time to be queried.
		--Also comment out 2 of the 3 pairs of variable setting in section ===SPECIFY EXTRACTION TYPE HERE===

		declare @cntr int, @batchCreatedDateTimeStart datetime, @stepNameFromBatchRecord varchar(50), @batchIsLive bit
		set nocount on

		Select @batchCreatedDateTimeStart = DATEADD(day, -8, CURRENT_TIMESTAMP);

		--===SPECIFY EXTRACTION TYPE HERE===
		--Duplicate EXTRACTION batch parameters. Use (comment out other 2) 1 of the 3 choices below to detect duplicates in either the MP96 extraction process or the EMag/EZMag one. 

		--For MagNa Pure 96 Extraction (ClinMicro) extraction option #1.
		--Use for looking at the step where test codes are gotten from soft and PendingReaction(s) are set. This will NOT show the current pending reactions, only the ones created based on the test codes retrieved.
		--select @stepNameFromBatchRecord = 'Get Test Code(s); Set PendingReaction(s)'
		--select @batchIsLive = 0

		--For MagNa Pure 96 Extraction (ClinMicro) extraction option #2.
		--Use for looking at the step where samples are made live and that holds the current PendingReaction(s) for the specimens.
		select @stepNameFromBatchRecord ='Confirm POS & NEG Extr. Cntrls passed Verification'
		select @batchIsLive = 1

		--For EMag/EZMag extracted samples onboarding process.
		--The step where test codes are retrieved, pending reactions set, and samples are made live
		--select @stepNameFromBatchRecord ='Get Test Code(s); Set PendingReaction(s)'
		--select @batchIsLive = 1

		--===END OF SPECIFY EXTRACTION TYPE HERE SECTION===


		--dupTempTable creation
		IF OBJECT_ID('tempdb..#dupTempTable') IS NOT NULL PRINT '#dupTempTable exists so dropping it.'
		IF OBJECT_ID('tempdb..#dupTempTable') IS NOT NULL DROP TABLE #dupTempTable

		CREATE TABLE #dupTempTable
		(
			specID VARCHAR(20),
			totalInstancesCount int,
			processInstName varchar(50),
			processDefName varchar(150),
			[biologics.id] BIGINT,
			subProcStepBatchId INT,
			batchName varchar(50),
			resevSuppPosition VARCHAR(5),
			testCodes varchar(150),
			pendingRxns varchar(150),
			repeatReason varchar(50),
			processingStatus varchar(50)
		)

		--===========
		--Get all unique patient samples (exclude controls) that exist given the criteria. All these are candidates for being dups but may not be.
		INSERT INTO #dupTempTable(specID, totalInstancesCount, processInstName, processDefName, [biologics.id], subProcStepBatchId, batchName, resevSuppPosition, testCodes, pendingRxns, repeatReason, processingStatus)
		SELECT distinct b.specid, 0, '', '', 0, 0, '','','','','',''
		FROM 
			dbo.Biologics b inner join dbo.SubProcessStepBatches spsb on b.subProcStepBatchId = spsb.id
			inner join SubProcessSteps sps on spsb.id = sps.subProcStepBatchIdEnd 
			inner join dbo.SubProcesses sp on sps.subProcId = sp.id
			inner join dbo.Processes p on sp.processId = p.id
		where
			spsb.creationDT>=@batchCreatedDateTimeStart and spsb.name=@stepNameFromBatchRecord and spsb.alive=@batchIsLive and b.specid <>''
			and not exists(select * from dbo.BiologicsAttributes ba where ba.biolsId=b.id and ba.biolsAttrDefId=11)
			and (SELECT COUNT(*) FROM  dbo.Processes AS p2 INNER JOIN dbo.SubProcesses AS sp2 ON p2.id = sp2.processId INNER JOIN dbo.SubProcessSteps AS sps2 ON sp2.id = sps2.subProcId
				WHERE (sps2.subProcStepDefId = 19) AND (p2.id = p.id)) <> 1
		order by b.specID desc

		select @cntr = count(*) from #dupTempTable
		print cast (@cntr as varchar (20)) + ' unique specimen id records exist with potential duplicates.'
		--select * from #dupTempTable order by specid desc

		--==========
		--determine and store if there are additional instances of each unique patient specimen id
		update #dupTempTable set 
				totalInstancesCount = (select count(b.id) FROM dbo.Biologics b inner join dbo.SubProcessStepBatches spsb on b.subProcStepBatchId = spsb.id
									inner join SubProcessSteps sps on spsb.id = sps.subProcStepBatchIdEnd 
									inner join dbo.SubProcesses sp on sps.subProcId = sp.id
									inner join dbo.Processes p on sp.processId = p.id
		where 
			spsb.creationDT >= @batchCreatedDateTimeStart and spsb.name = @stepNameFromBatchRecord and spsb.alive = @batchIsLive and b.specID = #dupTempTable.specID
			and (SELECT COUNT(*) FROM  dbo.Processes AS p2 INNER JOIN dbo.SubProcesses AS sp2 ON p2.id = sp2.processId INNER JOIN dbo.SubProcessSteps AS sps2 ON sp2.id = sps2.subProcId
		    WHERE (sps2.subProcStepDefId = 19) AND (p2.id = p.id)) <> 1)

		--==========
		--remove all the records not having > 1 instance of existence given the criteria (i.e. remove the duplicates)
		delete #dupTempTable where totalInstancesCount < 2
		--select * from #dupTempTable order by specid desc

		--==========
		-- add specific process, batch, and biologics info for the first instance of each patient specimen id
		update #dupTempTable set 
			[biologics.id] = b.id, subProcStepBatchId = spsb.id, batchName=spsb.name, resevSuppPosition=b.resevSuppPosition, processInstName = p.label, 
			processDefName = (select pd.title from dbo.ProcessDefs pd where pd.id = p.processDefId),
			testCodes = (select value from dbo.BiologicsAttributes ba where ba.biolsId=b.id and ba.biolsAttrDefId=18),
			pendingRxns = (select value from dbo.BiologicsAttributes ba where ba.biolsId=b.id and ba.biolsAttrDefId=19),
			repeatReason = (select value from dbo.BiologicsAttributes ba where ba.biolsId=b.id and ba.biolsAttrDefId=105),
			processingStatus = (select value from dbo.BiologicsAttributes ba where ba.biolsId=b.id and ba.biolsAttrDefId=88)
		from 
			dbo.Biologics b inner join dbo.SubProcessStepBatches spsb on b.subProcStepBatchId = spsb.id inner join SubProcessSteps sps on sps.subProcStepBatchIdEnd = spsb.id inner join dbo.SubProcesses sp on sp.id=sps.subProcId inner join dbo.Processes p on p.id = sp.processId
		where 
			spsb.creationDT >= @batchCreatedDateTimeStart and spsb.name = @stepNameFromBatchRecord and spsb.alive = @batchIsLive and b.specID = #dupTempTable.specID

		select @cntr = count(*) from #dupTempTable
		print cast (@cntr as varchar (20)) + ' specimens with at least one other instance.'
		--select * from #dupTempTable order by totalInstancesCount desc, specID

		--==========
		--now get all the other instances in addition to the single one identified so far (but not getting it again).
		insert into #dupTempTable ([biologics.id], subProcStepBatchId, batchName, specID, resevSuppPosition, totalInstancesCount, processInstName, processDefName, testCodes, pendingRxns, repeatReason, processingStatus)
		SELECT 
			b.id, spsb.id, spsb.name, b.specID, b.resevSuppPosition, 0, p.label, (select pd.title from dbo.ProcessDefs pd where pd.id = p.processDefId),
			testCodes = (select value from dbo.BiologicsAttributes ba where ba.biolsId=b.id and ba.biolsAttrDefId=18),
			pendingRxns = (select value from dbo.BiologicsAttributes ba where ba.biolsId=b.id and ba.biolsAttrDefId=19),
			repeatReason = (select value from dbo.BiologicsAttributes ba where ba.biolsId=b.id and ba.biolsAttrDefId=105),
			processingStatus = (select value from dbo.BiologicsAttributes ba where ba.biolsId=b.id and ba.biolsAttrDefId=88)
		FROM 
			dbo.Biologics b inner join dbo.SubProcessStepBatches spsb on b.subProcStepBatchId = spsb.id 
			inner join SubProcessSteps sps on spsb.id = sps.subProcStepBatchIdEnd 
			inner join dbo.SubProcesses sp on sps.subProcId = sp.id
			inner join dbo.Processes p on sp.processId = p.id
		where 
			spsb.creationDT >= @batchCreatedDateTimeStart and spsb.name = @stepNameFromBatchRecord and spsb.alive = @batchIsLive 
			and b.id not in (select [biologics.id] from #dupTempTable) and b.specID in (select specID from #dupTempTable)
				
		--select * from #dupTempTable order by specID, totalInstancesCount desc, processInstName, resevSuppPosition


		--==========
		--Final display of just dups and the basic process/batch/specimen info about them.
		select @cntr = count(*) from #dupTempTable
		print cast (@cntr as varchar (20)) + ' total specimens existing more than once each.'
		select specID,processInstName, processDefName, resevSuppPosition, processingStatus, testCodes, pendingRxns, repeatReason
		from #dupTempTable
		order by specID, processInstName, resevSuppPosition";

		try
		{
			var results = conLabOps.QueryAsync<Duplicates>(duplicatesQuery).Result.ToList();

			return results;

		}
		catch (Exception e)
		{
			Console.WriteLine(e);
			throw;
		}
	}

	public static IEnumerable<PendingSamples> GetPendingSamples()
	{
		var appSettings = new Startup();
		var conString = appSettings.Config?.GetSection("ConnectionStrings").GetValue<string>("windows-auth-PROD");

		using var conLabOps = new SqlConnection(conString);
		if (conLabOps.State == ConnectionState.Closed) conLabOps.Open();

		const string pendingSamplesQuery =
			@"-- Proposed enhanced version based on Brian's feedback here: https://dev.azure.com/mclm/GBS%20CAD/_workitems/edit/1723897

			declare @pending datetime
			Select @pending = DATEADD(day, -7, CURRENT_TIMESTAMP);

			SELECT distinct
				b.specID, pd.title, p.label as Batch,
				(Select top 1 value from dbo.BiologicsAttributes ba2 where ba2.biolsAttrDefId = 18 and ba2.biolsId = b.id) as testsOrdered,
				 (SELECT top 1 spsbja.creationDT 	FROM  dbo.Processes AS pja 
					INNER JOIN dbo.SubProcesses AS spja ON pja.id = spja.processId 
					INNER JOIN dbo.SubProcessSteps AS spsja ON spsja.subProcId = spja.id 
					INNER JOIN dbo.SubProcessStepDefs spsdja on spsdja.id = spsja.subProcStepDefId
					INNER JOIN dbo.SubProcessStepBatches AS spsbja ON spsja.subProcStepBatchIdEnd = spsbja.id 
					INNER JOIN dbo.Biologics AS bja ON spsbja.id = bja.subProcStepBatchId 
					INNER JOIN dbo.ProcessDefs AS pdja ON pja.processDefId = pdja.id 
				WHERE  pdja.categoryTypeDefId = 12 and pja.beginDT >= @pending -- and spsja.endDT is not null
					and bja.specID = b.specID
					and spsdja.actionText like 'Janus A%' Order by spsbja.creationDT asc) as janusA,
				 (SELECT top 1 spsbjb.creationDT 	FROM  dbo.Processes AS pjb 
					INNER JOIN dbo.SubProcesses AS spjb ON pjb.id = spjb.processId 
					INNER JOIN dbo.SubProcessSteps AS spsjb ON spsjb.subProcId = spjb.id 
					INNER JOIN dbo.SubProcessStepDefs spsdjb on spsdjb.id = spsjb.subProcStepDefId
					INNER JOIN dbo.SubProcessStepBatches AS spsbjb ON spsjb.subProcStepBatchIdEnd = spsbjb.id 
					INNER JOIN dbo.Biologics AS bjb ON spsbjb.id = bjb.subProcStepBatchId 
					INNER JOIN dbo.ProcessDefs AS pdjb ON pjb.processDefId = pdjb.id 
				WHERE  pdjb.categoryTypeDefId = 12 and pjb.beginDT >= @pending -- and spsjb.endDT is not null
					and bjb.specID = b.specID
					and spsdjb.actionText like 'Confirm Eluate Transfer%' Order by spsbjb.creationDT asc) as janusB,
				 (SELECT top 1 spsbjc.creationDT 	FROM  dbo.Processes AS pjc 
					INNER JOIN dbo.SubProcesses AS spjc ON pjc.id = spjc.processId 
					INNER JOIN dbo.SubProcessSteps AS spsjc ON spsjc.subProcId = spjc.id 
					INNER JOIN dbo.SubProcessStepDefs spsdjc on spsdjc.id = spsjc.subProcStepDefId
					INNER JOIN dbo.SubProcessStepBatches AS spsbjc ON spsjc.subProcStepBatchIdEnd = spsbjc.id 
					INNER JOIN dbo.Biologics AS bjc ON spsbjc.id = bjc.subProcStepBatchId 
					INNER JOIN dbo.ProcessDefs AS pdjc ON pjc.processDefId = pdjc.id 
				WHERE  pdjc.categoryTypeDefId in (13,14) and pjc.beginDT >= @pending -- and spsjc.endDT is not null
					and bjc.specID = b.specID
					and spsdjc.actionText like 'Create Janus .csv and LC%' Order by spsbjc.creationDT asc) as janusC,
				(SELECT top 1 spsbalz.creationDT 	FROM  dbo.Processes AS palz 
					INNER JOIN dbo.SubProcesses AS spalz ON palz.id = spalz.processId 
					INNER JOIN dbo.SubProcessSteps AS spsalz ON spsalz.subProcId = spalz.id 
					INNER JOIN dbo.SubProcessStepDefs spsdalz on spsdalz.id = spsalz.subProcStepDefId
					INNER JOIN dbo.SubProcessStepBatches AS spsbalz ON spsalz.subProcStepBatchIdEnd = spsbalz.id 
					INNER JOIN dbo.Biologics AS balz ON spsbalz.id = balz.subProcStepBatchId 
					INNER JOIN dbo.ProcessDefs AS pdalz ON palz.processDefId = pdalz.id 
				WHERE  pdalz.categoryTypeDefId in (13,14) and palz.beginDT >= @pending -- and spsalz.endDT is not null
					and balz.specID = b.specID
					and spsdalz.actionText like 'Parse results from .i%' Order by spsbalz.creationDT asc) as analysis,
				(SELECT top 1 spsbver.creationDT 	FROM  dbo.Processes AS pver 
					INNER JOIN dbo.SubProcesses AS spver ON pver.id = spver.processId 
					INNER JOIN dbo.SubProcessSteps AS spsver ON spsver.subProcId = spver.id 
					INNER JOIN dbo.SubProcessStepDefs spsdver on spsdver.id = spsver.subProcStepDefId
					INNER JOIN dbo.SubProcessStepBatches AS spsbver ON spsver.subProcStepBatchIdEnd = spsbver.id 
					INNER JOIN dbo.Biologics AS bver ON spsbver.id = bver.subProcStepBatchId 
					INNER JOIN dbo.ProcessDefs AS pdver ON pver.processDefId = pdver.id 
				WHERE  pdver.categoryTypeDefId in (13,14) and pver.beginDT >= @pending -- and spsver.endDT is not null
					and bver.specID = b.specID
					and spsdver.actionText like 'Click ''Verify'' %' Order by spsbver.creationDT asc) as verifying,
				(Select top 1 value from dbo.BiologicsAttributes ba3 where ba3.biolsAttrDefId = 19 and ba3.biolsId = b.id) as testsPending,
				( SELECT max(plc.endDt) 
						FROM     dbo.Processes AS plc INNER JOIN dbo.SubProcesses AS splc ON plc.id = splc.processId 
							INNER JOIN dbo.SubProcessSteps AS spslc ON spslc.subProcId = splc.id 
							INNER JOIN dbo.SubProcessStepBatches AS spsblc ON spslc.subProcStepBatchIdEnd = spsblc.id 
							INNER JOIN dbo.Biologics AS blc ON spsblc.id = blc.subProcStepBatchId INNER JOIN dbo.ProcessDefs AS pdlc ON plc.processDefId = pdlc.id 
						WHERE  pdlc.categoryTypeDefId in (13,14) and plc.endDT > @pending and blc.specID = b.specID) as allCompletePendingBatches
			FROM     dbo.Processes AS p INNER JOIN
			                  dbo.SubProcesses AS sp ON p.id = sp.processId INNER JOIN
			                  dbo.SubProcessSteps AS sps ON sps.subProcId = sp.id INNER JOIN
			                  dbo.SubProcessStepBatches AS spsb ON sps.subProcStepBatchIdEnd = spsb.id INNER JOIN
			                  dbo.Biologics AS b ON spsb.id = b.subProcStepBatchId INNER JOIN
			                  dbo.ProcessDefs AS pd ON p.processDefId = pd.id 
			WHERE  pd.categoryTypeDefId in (12) --,13,14)--12 = IP, 13 = LC480, 14 = LC 2.0
			and not exists (select 1 from BiologicsAttributes ba where ba.biolsAttrDefId=11 and biolsid = b.id) --don't return controls (IsControl = def id 11)
			and b.specID <> '' and spsb.alive = 1 and pd.title like 'MagNa Pure%'
			and p.beginDT >= @pending
			order by p.label desc, b.specID asc, pd.title;";
		
		try
		{
			var results = conLabOps.QueryAsync<PendingSamples>(pendingSamplesQuery).Result.ToList();
			return results;

		}
		catch (Exception e)
		{
			Console.WriteLine(e);
			throw;
		}
		
		    
	}

}