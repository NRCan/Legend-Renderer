using System.Collections.Generic;

namespace GSCLegendRendererPro.Utilities
{
    public class Constants
    {
        #region Project Database related constants

        public class Database
        {

            public const double CurrentDBVersion = 4.0; //Update this when DB schema is changed
            public const double DBVersion_300 = 3.0; //Never implemented, skipped from 2.10 to 4.0 to align with software versioning
            public const double DBVersion_210 = 2.10; 

            //Tables
            public const string TParticipant = "P_PARTICIPANT";
            public const string TPerson = "P_PERSON";
            public const string TLegendDescription = "P_LEGEND_DESCRIPTION"; //Deprecated table
            public const string TStudyAreaIndex = "P_STUDY_AREA_INDEX"; //Deprecated table
            public const string TSource = "P_SOURCE";
            public const string TOrganisation = "P_ORGANIZATION";
            public const string TLegendGene = "P_LEGEND";
            public const string TLegendTree = "P_LEGEND_INDEX"; //Deprecated table

            public const string TProject = "M_PROJECT";
            public const string TMActivity = "M_ACTIVITY";
            public const string TMetadata = "M_METADATA";
            public const string TSActivity = "M_SUB_ACTIVITY";

            public const string TGeolineSymbol = "SYMBOL_GEOLINES";
            public const string TGeopointSymbol = "SYMBOL_GEOPOINTS";

            public const string TExtenAttrb = "CARTOGRAPHIC_";//Last part will depend on user choice of theme name

            public const string TDomainAssigned = "DOMAINS_ASSIGNED";//Deprecated table
            public const string TDomainDID = "DOMAINS_DID"; //Deprecated table
            public const string TDomainPID = "DOMAINS_PID";//Deprecated table
            public const string TDomainSID = "DOMAINS_SID";//Deprecated table
            public const string TSubtypes = "M_SUBTYPES";//Deprecated table

            public const string TGeoEvent = "P_GEO_EVENT";

            //Features
            public const string FGeoline = "GEO_LINES";
            public const string FLabel = "LABELS";
            public const string FGeopoly = "GEO_POLYS";
            public const string FStation = "F_STATION";//Deprecated table
            public const string FStudyArea = "P_STUDY_AREA";
            public const string FStudyAreaAlias = "Study Areas"; //GIS display of feature class name
            public const string FCGMIndex = "P_CGM";
            public const string FGeopoint = "GEO_POINTS";
            public const string FCartoPoint = "CARTOGRAPHIC_POINTS";

            //Feature datasets
            public const string FDField = "FIELD_OBS"; //Ganfeld
            public const string FDGeo = "GEO";

            //Topographic database
            public const string topoDBName = "Project_Elements";

            //Ganfeld tables
            public const string gEarthMath = "F_EARTHMAT";//Deprecated table
            public const string gMA = "F_MA";//Deprecated table
            public const string gMetadata = "F_METADATA";//Deprecated table
            public const string gMineral = "F_MINERAL";//Deprecated table
            public const string gSample = "F_SAMPLE";//Deprecated table
            public const string gStruc = "F_STRUC";//Deprecated table
            public const string gPhoto = "F_PHOTO";//Deprecated table
            public const string gEnviron = "F_ENVIRON";//Deprecated table
            public const string gPFlow = "F_PFLOW";//Deprecated table
            public const string gSoil = "F_SOILPRO";//Deprecated table
            public const string gBiogeo = "F_BIOGEO";//Deprecated table

            //Ganfeld feature
            public const string gFCLinework = "F_LINEWORK";//Deprecated table
            public const string gFCStation = "F_STATION";//Deprecated table
            public const string gFCTraverses = "F_TRAVERSE";//Deprecated table

            //Old names (to be kept for legacy between DB versions)
            public const string TOrganisation_160915 = "P_ORGANISATION";
            public const string TGeoEvent_160915 = "GEO_EVENT";
            public const string TSActivity_141003 = "S_ACTIVITY";
            public const string FLabel_141003 = "P_LABELS";
            public const string TExtenAttrb_141003 = "EXTENDED_ATTRIBUTES";
            public const string TExtenAttrb_160322 = "CARTOGRAPHIC_ATTRIBUTES";
            public const string TGeolineSymbol_141003 = "GEOLINE_SYMBOLS";
            public const string TGeopointSymbol_141003 = "GEOPOINT_SYMBOLS";
            public const string TLegendTree_141003 = "LEGEND_TREETABLE";
            public const string TLegendGene_160411 = "P_LEGEND_GENERATOR";
            public const string TLegendTree_160411 = "P_LEGEND_TREETABLE";
            public const string FCGMIndex_160411 = "P_CGM_INDEXMAP";

            //Topology
            public const string Topology = "GEO_TOPO";

            //Temporary names for features and tables
            public const string tLegendGeneratorTemp = "P_LEGEND_VIEW";
            public const string tLegendGeneratorTemp191224 = "P_LEGEND_GENERATOR";
            //Relationship
            public const string rel_tLegend_fGeopoly = "Rel_LegendGenerator_GeoPolys";
            public const string rel_prefix_CartoPnt = "Rel_CartoPnt_";

        }

        public static class DatabaseFields
        {
            //Main
            public const string ObjectID = "OBJECTID";
            public const string SourceID = "SOURCEID";

            #region TABLES 

            #region Participant Table
            public const string ParticipantRoleDesc = "ROLEDESC";
            public const string ParticipantPersonID = "PERSONID";
            //public const string ParticipantActivityID = "ACTIVITYID";
            public const string ParticipantRole = "PARTROLE";
            public const string ParticipantGeolCode = "GEOLCODE";
            public const string ParticipantStartDate = "STARTDATE";
            public const string ParticipantRemarks = "REMARKS";
            public const string ParticipantMetaID = "METAID";
            public const string ParticipantEndDate = "ENDDATE";
            public const string ParticipantMActID = "M_ACTIVITYID";
            public const string ParticipantSActID = "S_ACTIVITYID";
            public const string ParticipantID = "PARTICIPANTID";

            #endregion

            #region Person table
            public const string PersonFirstName = "FNAME";
            public const string PersonMiddleName = "MNAME";
            public const string PersonLastName = "LNAME";
            public const string PersonAbbr = "ABBREVNAME";
            public const string PersonPhone = "PHONE";
            public const string PersonEmail = "EMAIL";
            public const string PersonOrg = "ORGID";
            public const string PersonID = "PERSONID";
            public const string PersonAlias = "ALIAS";
            #endregion

            #region Sub Activity table
            public const string SubActivityName = "S_ACTIVITYNAME";
            public const string SubActivityID = "S_ACTIVITYID";
            public const string SubActivityMainID = "M_ACTIVITYID";
            public const string SubActivityAbbr = "ABBREVIATION";
            public const string SubActivityStart = "STARTDATE";
            public const string SubActivityEnd = "ENDDATE";
            public const string SubActivityDesc = "DESCRIPTION";
            #endregion

            #region Main Activity table
            public const string MainActivityName = "M_ACTIVITYNAME";
            public const string MainActID = "M_ACTIVITYID";
            public const string MainActProjectID = "PROJECTID";
            public const string MainActAbbr = "ABBREVIATION";
            public const string MainActStart = "STARTDATE";
            public const string MainActEnd = "ENDDATE";
            public const string MainActDesc = "DESCRIPTION";
            #endregion

            #region Organisation table
            public const string OrganisationName = "ORGNAME";
            public const string OrganisationID = "ORGID";
            public const string OrganisationAddress = "ORGADD";
            public const string OrganisationPhone = "ORGPHONE";
            public const string OrganisationEmail = "ORGEMAIL";
            public const string OrganisationWeb = "ORGWWW";
            public const string OrganisationAbbr = "ORGABBREV";
            #endregion

            #region Project table
            public const string ProjectID = "PROJECTID";
            public const string ProjectName = "PROJECTNAME";
            public const string ProjectNom = "PROJECTNOM";
            public const string ProjectAbbr = "PROJECTABBREV";
            public const string ProjectStart = "STARTDATE";
            public const string ProjectEnd = "ENDDATE";
            public const string ProjectRemarks = "REMARKS";
            public const string ProjectWebLink = "WEBLINK";
            public const string ProjectCode = "PROJECTCODE";
            #endregion

            #region Master Geoline Symbols
            public const string MGeolineSelectCode = "SelectCode"; //Used to get symbols into templates within Arc Map
            public const string MGeolineLegendDescription = "Legend_Description";

            public const string MGeolineID = "GEOLINEID";
            public const string MGeolineFGDC = "GSC_SYMBOL";
            #endregion

            #region Legend generator table
            public const string LegendLabelID = "LEGENDITEMID";
            public const string LegendGISDisplay = "GIS_DISPLAY_NAME";
            public const string LegendItemType = "ELEMENT"; //Version 4.0
            public const string LegendSymbol = "STYLE1"; //Version 4.0
            public const string LegendSymbol2 = "STYLE2"; //Version 4.0
            public const string LegendLabel1 = "LABEL1"; //Version 4.0
            public const string LegendLabel1Style = "LABEL1STYLE"; //Version 4.0
            public const string LegendLabel2 = "LABEL2"; //Version 4.0
            public const string LegendLabel2Style = "LABEL2STYLE"; //Version 4.0
            public const string LegendHeading = "HEADING"; //Version 4.0
            public const string LegendColumn = "ELEMENT_COLUMN"; //Version 4.0
            public const string LegendOrder = "ELEMENT_ORDER"; //Version 4.0
            public const string LegendDescription = "DESCRIPTION"; //Version 4.0
            public const string LegendGeolRank = "GEOLRANK";
            public const string LegendOverprint = "OVERPRINT"; //Version 4.0

            public const string LegendSymbol_190101 = "GSC_SYMBOL"; //Version 2.10
            public const string LegendSymType_190101 = "SYM_TYPE"; //Version 2.10
            public const string LegendLabelName_190101 = "NAME"; //Version 2.10
            public const string LegendMapUnit_190101 = "MAPUNIT"; //Version 2.10
            public const string LegendAnnotation_190101 = "ANNOTATION"; //Version 2.10
            public const string LegendOrder_190101 = "LEGEND_ORD"; //Version 2.10
            public const string LegendIndentation_190101 = "INDENT"; //Version 2.10
            public const string LegendItemType_190101 = "LEGEND_ITEMTYPE"; //Version 2.10
            #endregion

            #region Legend description table
            public const string LegendDescription_190101 = "DESCRIPTION"; //Version 2.10
            public const string LegendDescriptionID_190101 = "LEGDESCRIPTIONID"; //Version 2.10
            #endregion

            #region Legend tree table
            public const string LegendTreeDescID = "LEGDESCRIPTIONID";
            public const string LegendTreeCGM = "CGM_MAPID";
            public const string LegendTreeItemID = "LEGENDITEMID";
            #endregion

            #region Table M_METADATA
            public const string MetadataPurposeFr = "PURPOSE_FR";
            public const string MetadataPurpose = "PURPOSE";
            public const string MetadataNAP = "NAP_METAID";
            #endregion

            #region Table study area index
            public const string TStudyAreaName = "TABLENAME";
            public const string TStudyAreaRowID = "TABLE_RELATEDID";
            public const string TStudyAreaFC = "FC_NAME";
            #endregion

            #region Table P_SOURCE
            public const string TSourceID = "SOURCEID";
            public const string TsourceName = "SOURCENAME";
            public const string TSourceRemarks = "REMARKS";
            public const string TSourceDOI = "DOI";
            public const string TSourceAbbr = "ABBREVIATION";
            public const string TSourceFilePath = "FILEPATH";
            public const string TSourceExtended = "EXTENDED_SOURCE";
            #endregion

            #region Table Geopoint_symbol
            public const string TGeopointID = "GEOPOINTID";
            public const string TGeopointLegendDesc = "Legend_Description";
            public const string TGeopointFGDC = "GSC_SYMBOL";
            public const string TGeopointSelectCode = "SelectCode";
            #endregion

            #region GEO_EVENT

            public const string TGeoEventID = "GEOEVENT_ID";
            public const string TGeoEventName = "GEOEVENT_NAME";
            public const string TGeoEventAgeMinPrefix = "AGEMIN_PREFIX";
            public const string TGeoEventAgeMinTimescale = "AGEMIN_TIMESCALE";
            public const string TGeoEventAgeMinValue = "AGEMIN_VALUE";
            public const string TGeoEventAgeMinCertainty = "AGEMIN_CERTAINTY";

            public const string TGeoEventAgeMaxPrefix = "AGEMAX_PREFIX";
            public const string TGeoEventAgeMaxTimescale = "AGEMAX_TIMESCALE";
            public const string TGeoEventAgeMaxValue = "AGEMAX_VALUE";
            public const string TGeoEventAgeMaxCertainty = "AGEMAX_CERTAINTY";

            public const string TGeoEventSourceID = "SOURCEID";

            #endregion

            #endregion

            #region FEATURES

            #region Feature Label
            public const string FLabelID = "LABELID";
            public const string FLabelIDAlias = "Label";
            public const string FLabelGeoEventID = "GEOEVENT_ID";
            public const string FLabelSymbol = "GSC_SYMBOL";
            #endregion

            #region Feature Geoline
            public const string FGeolineSubtype = "GEOLINETYPE";
            public const string FGeolineQualif = "QUALIFIER";
            public const string FGeolineConf = "CONFIDENCE";
            public const string FGeolineAtt = "ATTITUDE";
            public const string FGeolineGeneration = "GENERATION";
            public const string FGeolineFGDC = "GSC_SYMBOL";
            public const string FGeolineID = "GEOLINEID";
            public const string FGeolineBoundary = "ISBOUNDARY";
            public const string FGeolineMovement = "MOVEMENT";
            public const string FGeolineHangwall = "HWALLDIR";
            public const string FGeolineFoldTrend = "FOLDTREND";
            public const string FGeolineArrowDir = "ARROWDIR";
            public const string FGeolineDisplayPub = "DISPLAYPUB";
            public const string FGeolineGeoEventID = "GEOEVENT_ID";
            public const string FGeolineName = "NAME";
            public const string FGeolineOriginalCode = "ORIGCODE";
            public const string FGeolineRemarks = "REMARKS";
            public const string FGeolineRemarksEdit = "EDITREMARKS";
            #endregion

            #region Feature Geopolys
            public const string FGeopolyLabel = "LABELID";
            public const string FGeopolyRemark = "REMARKS";
            public const string FGeopolyFGDC = "GSC_SYMBOL";
            #endregion

            #region Feature Study area
            public const string FStudyAreaAbbr = "ABBREVIATION";
            public const string FStudyAreaRemarks = "REMARKS";
            public const string FStudyAreaEast = "EAST_EXTENT_COORD";
            public const string FStudyAreaWest = "WEST_EXTENT_COORD";
            public const string FStudyAreaNorth = "NORTH_EXTENT_COORD";
            public const string FStudyAreaSouth = "SOUTH_EXTENT_COORD";
            public const string FStudyAreaRelatedID = "TABLE_RELATEDID";
            #endregion

            #region Feature CGM maps
            public const string FCGM_ID = "CGM_MAPID";
            public const string FCGM_Name = "MAPNAME";
            public const string FCGM_Abstract = "ABSTRACT";
            public const string FCGM_Resume = "RESUME";
            public const string FCGM_DescNote = "DESCNOTE";
            public const string FCGM_NapID = "NAP_METAID";
            public const string FCGM_RelatedID = "TABLE_RELATEDID";
            public const string FCGM_East = "EAST_EXTENT_COORD";
            public const string FCGM_West = "WEST_EXTENT_COORD";
            public const string FCGM_North = "NORTH_EXTENT_COORD";
            public const string FCGM_South = "SOUTH_EXTENT_COORD";
            public const string FCGM_Remarks = "REMARKS";
            //public const string FCGM_Abbr = "ABBREVIATION";
            #endregion

            #region Feature F_STATION
            public const string FStationEasting = "EASTING";
            public const string FStationNorthing = "NORTHING";
            public const string FStationLat = "LATITUDE";
            public const string FStationLong = "LONGITUDE";
            public const string FStationID = "STATIONID";
            public const string FStationObjectID = "OBJECTID";
            #endregion

            #region Feature geopoint
            public const string FGeopointID = "GEOPOINTID";
            public const string FGeopointType = "GEOPOINTTYPE";
            public const string FGeopointSubset = "GEOPOINTSUBSET";
            public const string FGeopointStrucAtt = "STRUCATTITUDE";
            public const string FGeopointStrucGene = "STRUCGENERATION";
            public const string FGeopointStrucYoung = "STRUCYOUNGING";
            public const string FGeopointStrucMethod = "STRUCMETHOD";
            public const string FGeopointFGDC = "GSC_SYMBOL";
            public const string FGeopointAzimuth = "AZIMUTH";
            public const string FGeopointFlat = "FLATTENING";
            public const string FGeopointStrain = "STRAIN";
            public const string FGeopointDipPlunge = "DIPPLUNGE";
            public const string FGeopointRelatedStruc = "RELATEDSTRUC";
            public const string FGeopointStrucID = "F_STRUCID";
            public const string FGeopointSenseEvid = "SENSE_EVID";
            public const string FGeopointRemark = "REMARKS";
            public const string FGeopointDipDescription = "DIPDESC";
            public const string FGeopointOriginalCode = "ORIGCODE";
            public const string FGeopointEditRemarks = "EDITREMARKS";
            public const string FGeopointDisplayFrom = "DISPLAY_FROM";
            public const string FGeopointDisplayTo = "DISPLAY_TO";
            public const string FGeopointDisplayPub= "DISPLAYPUB";

            #endregion

            #region Carto Points

            public const string FCartoPointID = "POINTID";
            public const string FCartoPointDisplayFrom = "DISPLAY_FROM";
            public const string FCartoPointSymbol = "SYMBOL";
            public const string FCartoPointAngle = "SYM_ANGLE";
            public const string FCartoPointEasting = "EASTING";
            public const string FCartoPointNorthing = "NORTHING";
            public const string FCartoPointLongitude = "LONGITUDE";
            public const string FCartoPointLatitude = "LATITUDE";
            public const string FCartoPointDatumZone = "DATUMZONE";
            public const string FCartoPointSourceID = "SOURCEID";
            public const string FCartoPointAltitude = "ALTITUDE";
            public const string FCartoPointTheme = "THEME";
            public const string FCartoPointLegendID = "LEGENDITEMID"; //Version 2.8

            #endregion

            #endregion

            #region Editor tracking related fields
            public const string ETCreatorID = "CREATORID";
            public const string ETEditorID = "EDITORID";
            public const string ETCreateDate = "CREATEDATE";
            public const string ETEditDate = "EDITDATE";
            #endregion

            #region Map Specific legends NOT IN DEFAULT DB SCHEMA
            public const string LegendTempIsVisible = "IS_INSIDE";
            public const string LegendTempIsVisibleAlias = "Is Inside Map";
            #endregion

            #region Old names (to be kept for legacy between DB versions)
            public const string FGeolineFGDC_151001 = "FGDC_SYMBOL";
            public const string LegendSymbol_151001 = "SYMBOL";
            public const string FGeopointFGDC_151001 = "FGDC_Symbol";
            public const string FGeopointFGDC_151015 = "FGDC_SYMBOL";
            public const string LegendEon_160330 = "EON";
            public const string LegendEra_160330 = "ERA";
            public const string LegendPeriod_160330 = "PERIOD";
            public const string FGeolineMinAge_160101 = "MINAGE";
            public const string FGeolineMaxAge_160101 = "MAXAGE";
            #endregion
        }

        public static class DatabaseDomains
        {
            //Participant role
            public const string ActivityRole = "ActivityRole_DID";

            //Boolean yes or no
            public const string BoolYesNo = "BooleanTruth_DID";

            //Map units
            //public const string Eon = "Eon_DID"; //Removed in version 2.7 of the schema
            //public const string Era = "Era_DID"; //Removed in version 2.7 of the schema
            //public const string Period = "Period_DID"; //Removed in version 2.7 of the schema
            public const string MapUnit = "MapUnit_PID";

            //Source
            public const string Source = "SourceRef_PID";

            //Organisation
            //public const string Org = "Organisation_PID"; //Removed in version 2.6 of the schema

            //Boundary for geolines
            public const string Bound = "Boundary_DID";

            //Geological rank
            public const string geolRank = "RankTerm_DID";

            //Geoline movement
            public const string geolMovement = "Fault_move_SID";

            //Participant 
            public const string participant = "Participant_PID";

            //Keyworks
            public const string projectKeyword = "PID";

            //Legend
            public const string legendSymbolType = "LegendSymbolType_DID";
            public const string legendSymbolTheme = "LegendItemTheme_PID";
            public const string legendOverprint = "OverprintLevel_DID";

            //Age
            public const string ageDesignator = "AgeDesignator_DID";
            public const string agePrefix = "AgePrefix_DID";

            //Carto theme
            public const string cartoTheme = "CartoTheme_PID";

            //Other
            public const string qualifLimitSID = "Qualif_limit_SID";
            public const string qualifFaultSID = "Qualif_fault_SID";
            public const string qualifOverprintSID = "Qualif_overprint_SID";
            public const string qualifShearSID = "Qualif_shear_SID";
            public const string qualifContactSID = "Qualif_contact_SID";
            public const string qualifConstructSID = "Qualif_construct_SID";
            public const string subsetLinearSID = "Subset_Linear_SID";
            public const string subsetPlanarSID = "Subset_Planar_SID";
            public const string strucGenerationSID = "Struc_Generation_SID";
            public const string strucAttitudeLinear = "Struc_Attitude_Linear_SID";
            public const string genUndefinedSID = "Gen_undefined_SID";
            public const string qualifLineamSID = "Qualif_lineam_SID";


        }

        public static class DatabaseDomainsValues
        {
            //BooleanTruth
            public const string BoolYes = "1";
            public const string BoolNo = "0";

            //Boundary for geolines
            public const string BoundYes = "02";
            public const string BoundNo = "01";

            //Confidence for geolines
            public const string ConfDef = "01";
            public const string ConfAprox = "02";
            public const string ConfInf = "03";
            public const string ConfConcealed = "04";
            public const string ConfNotApp = "99";

            //Symbol type
            public const string SymTypeFill = "F";
            public const string SymTypeLine = "L";
            public const string SymTypePoint = "M";
            public const string SymTypeHeader1 = "H1";
            public const string SymTypeHeader2 = "H2";
            public const string SymTypeHeader3 = "H3";

            //Legend item type
            public const string legendItemMapUnit = "mapUnit";
            public const string legendItemGeoline = "geoline";
            public const string legendItemGeopoint = "geopoint";
            public const string legendItemHeader = "header";
            public const string legendItemField = "fieldPoint";
            public const string legendItemCartoPoint = "cartographicPoint";

            //geoline types
            public const string GeolineContacts = "10";
            public const string GeolineFaults = "11";
            public const string GeolineStructureLineaments = "13";
            public const string GeolineTrace = "17";

            //geoline qualifier
            public const string geolineDefineUnconformable = "1004";
            public const string geolineFaultNormal = "2001";

            //Geoline movement
            public const string geolineMovementN = "01";
            public const string geolineMovementNE = "02";
            public const string geolineMovementE = "03";
            public const string geolineMovementSE = "04";
            public const string geolineMovementS = "05";
            public const string geolineMovementSW = "06";
            public const string geolineMovementW = "07";
            public const string geolineMovementNW = "08";
            public const string geolineMovementUndef = "88";
            public const string geolineMovementNA = "99";

            //Geoline hangwall codes
            public const string geolineHgwallN = "01";
            public const string geolineHgwallNE = "02";
            public const string geolineHgwallE = "03";
            public const string geolineHgwallSE = "04";
            public const string geolineHgwallS = "05";
            public const string geolineHgwallSW = "06";
            public const string geolineHgwallW = "07";
            public const string geolineHgwallNW = "08";
            public const string geolineHgwallUndef = "88";
            public const string geolineHgwallNA = "99";

            //Geopoint subtype class
            public const int geopointPlanar = 1;
            public const int geopointLinear = 2;

            //Geopoint attitude 
            public const string geopointAttHoriUp = "07";
            public const string geopointAttInclinedUp = "01";
            public const string geopointAttHoriOverturned = "08";
            public const string geopointAttInclinedOverLess180 = "02";
            public const string geopointAttInclinedOverHigher180 = "03";
            public const string geopointAttUndef = "88";
            public const string geopointAttLinearPlunging = "10";
            public const string geopointAttLinearHori = "05";
            public const string geopointAttVerti = "04";

            //Not applicable
            public const string notAppicable = "99";

            //Carto Point themes
            public const string cartoPointThemeField = "fieldLegacy";

            //Eon values
            public const string eonPhanerozoic = "phanerozoic";

        }

        public static class DatabaseSubtypes
        {
            //Subtype codes For GEO_LINES
            public const string FGeolineSubContact = "10";
            public const string FGeolineSubThinUnit = "15";
            public const string FGeolineSubUnitConstruct = "16";
            public const string FGeolineSubFault = "11";
            public const string FGeolineSubOverprint = "19";
        }

        #endregion

        #region Project related constants

        public static class ProjectLayers
        {
            //Relative paths
            public const string MapUnitLayerRelPath = "\\Data\\LYR\\BEDROCK\\INTERPRETATION_MapUnits.lyr";

            //Default layer names
            public const string geoline = "Geolines";
            public const string fStation = "F_STATION";
            public const string cgmMapIndex = "CGM Map Index";
            public const string geopoint = "GeoPoints";
            public const string geopoly = "Map Units";
            public const string label = "Labels";
            public const string geolineTopology = "Geolines Topology";

            //Field layer names
            public const string earthmatDetail = "EARTHMAT_DETAIL";
            public const string earthmatGroup = "EARTHMAT_GROUP";
            public const string earthmatType = "EARTHMAT_TYPE";
            public const string sampleGeochemistry = "SAMPLE_Geochemistry";
            public const string sampleGeochrone = "SAMPLE_Geochronology";
            public const string sampleMineralogy = "SAMPLE_Mineralogy";
            public const string samplePT = "SAMPLE_PT_estimate";
            public const string sampleRepLitho = "SAMPLE_Representative_Lithology";
            public const string sampleThinSection = "SAMPLE_thin_section";
            public const string sampleAll = "SAMPLES_ALL";
            public const string sampleOther = "SAMPLE_Other";
            public const string photoAll = "PHOTOS_ALL";
            public const string strucAll = "STRUC_ALL";
            public const string strucLinearAll = "STRUC_LINEAR_ALL";
            public const string strucPlanDip_0 = "STRUC_PLAN_DIP_0";
            public const string strucPlanDip_0_90 = "STRUC_PLAN_DIP_0_90";
            public const string strucPlanDip_90 = "STRUC_PLAN_DIP_90";
            public const string interpretationGeoline = "INTERPRETATION_Geolines";


            //Name for theme layer
            public const string overprintThematic = "Overprints";
            public const string noOverprintThematic = "No Overprints";

            //New names for temp layers
            public const string Keyword = "Validation"; //Will also be used to find layers or not find layers in TOC.

            //Labels
            public const string labelLayerQC_RedMultipleLabel = "Labels " + Keyword + ": Problematic labels in same map units";

            //Map units 
            public const string MapUnitLayerQC_Red = "Map Unit " + Keyword + ": Adjacent polygons";
            public const string MapUnitLayerQC_Yellow = "Map Unit " + Keyword + ": Isolated polygons";
            public const string MapUnitLayerQC_RedMultipleLabel = "Map Unit " + Keyword + ": Contains different labels";
            public const string MapUnitLayerQC_YellowMultipleLabel = "Map Unit " + Keyword + ": Contains no labels";
            public const string MapUnitLayerQC_YellowSmallFeature1 = "Map Unit " + Keyword + ": Small features (Between 1 and 1.5 mm on paper map)";
            public const string MapUnitLayerQC_YellowSmallFeature2 = "Map Unit " + Keyword + ": Small features (Less than 1 mm on paper map)";
            public const string MapUnitLayerQC_YellowSmallFeature3 = "Map Unit " + Keyword + ": Small features without any label, less than 1.5 mm";

            //Geoline
            public const string GeolineLayerQC_BlueIsBoundary = "Geolines " + Keyword + ": IsBoundary = Yes";
            public const string GeolineLayerQC_PinkIsBoundary = "Geolines " + Keyword + ": IsBoundary = No";
            public const string GeolineLayerQC_DefinedConfidence = "Geolines " + Keyword + ": Confidence = Defined";
            public const string GeolineLayerQC_ApproxConfidence = "Geolines " + Keyword + ": Confidence = Approximate";
            public const string GeolineLayerQC_InferredConfidence = "Geolines " + Keyword + ": Confidence = Inferred";
            public const string GeolineLayerQC_ConcealeadConfidence = "Geolines " + Keyword + ": Confidence = Concealed";
            public const string GeolineLayerQC_NotApplicableConfidence = "Geolines " + Keyword + ": Confidence = Not Applicable";
            public const string GeolineLayerQC_ContactAgeMustBeFlip = "Geolines " + Keyword + ": Contact vs Age = Must be flip";
            public const string GeolineLayerQC_ContactAgeNoNeighbour = "Geolines " + Keyword + ": Contact vs Age = No neighbors or Undefined Legend Order";
            public const string GeolineLayerQC_MovementFaultFlip = "Geolines " + Keyword + ": Hangwall value should be flipped";
            public const string GeolineLayerQC_MovementFaultCode01 = "Geolines " + Keyword + ": Hangwall code value should be ";
            public const string GeolineLayerQC_MovementFaultCode02 = "Geolines " + Keyword + ": Hangwall code value should be ";
            public const string GeolineLayerQC_MovementFaultCode03 = "Geolines " + Keyword + ": Hangwall code value should be ";
            public const string GeolineLayerQC_MovementFaultCode04 = "Geolines " + Keyword + ": Hangwall code value should be ";
            public const string GeolineLayerQC_MovementFaultCode05 = "Geolines " + Keyword + ": Hangwall code value should be ";
            public const string GeolineLayerQC_MovementFaultCode06 = "Geolines " + Keyword + ": Hangwall code value should be ";
            public const string GeolineLayerQC_MovementFaultCode07 = "Geolines " + Keyword + ": Hangwall code value should be ";
            public const string GeolineLayerQC_MovementFaultCode08 = "Geolines " + Keyword + ": Hangwall code value should be ";
            public const string GeolineLayerQC_MovementFaultMissing = "Geolines " + Keyword + ": Movement is either 'Undefined' or 'Not Applicable' (" + DatabaseDomainsValues.geolineMovementUndef + ", " + DatabaseDomainsValues.geolineMovementNA + ")";
            public const string GeolineLayerQC_MovementFaultChangeSym = "Geolines " + Keyword + ": Verify symbol code";
            public const string GeolineLayerQC_ConcealedOverprintFullyInside = "Geolines " + Keyword + ": Concealed lines fully underneath overprints";
            public const string GeolineLayerQC_ConcealedOverprintTouches = "Geolines " + Keyword + ": Concealed lines that touches overprints";

            //Field Station
            public const string FStationLayerQC_BadEastNorth = "Field Station " + Keyword + ": Easting/Northing Mismatch";
            public const string FStationLayerQC_BadLatLong = "Field Station " + Keyword + ": Latitude/Longitude Mismatch";

            //Group layers
            public const string GroupInterpretation = "INTERPRETATION";

            //Group layer CGM
            public const string CGMGeologyBedrock = "BEDROCK";
            public const string CGMGeopoint = "GEOLOGY POINTS";
            public const string CGMGeoline = "GEOLOGY LINES";
            public const string CGMGeopoly = "GEOLOGY POLYGONS";
            public const string CGMOverprint = "GEOLOGY OVERPRINTS";
            public const string CGMGeologyType = "<GEOLOGY TYPE>";
            public const string CGMLimit = "CGM MAP LIMIT";

            //Default layer names CGM


        }

        public static class Symbol4Layers
        {
            //Label name for symbols
            public const string geolineLabelFieldName = "Legend Description";
            public const string geopolyLabelFieldName = "Legend Description";
            public const string geopointLabelFieldName = "Legend Description";

            //Default color if user hasn't entered any for map units
            public const string mapUnitDefaultColor = "2.04.01.011";

            //Field delimeter for multiple field symbolization
            public const string fieldDelimeter = ", ";

        }

        public static class Folders
        {
            //A list of folder to be created within a project main working folder
            public const string imageryFolder = "Imagery";
            public const string sourceFolder = "Source";
            public const string fieldDataFolder = "FieldData";
            public const string mxdFolder = "Templates";
            public const string styleFolder = "Style";

            //For publication
            public const string publicationFolder = "Dissemination"; //Parent
            public const string publicationDataFolder = "Data";
            public const string publicationShapeFolder = "SHP";
            public const string publicationBedrockFolder = "Bedrock";
            public const string publicationSurficialFolder = "Surficial";
            public const string publicationCartoElementsFolder = "CartoElements";
            public const string publicationDBFolder = "GDB";
            public const string publicatonXMLFolder = "XML";
            public const string publicationDataModelInfoFolder = "DataModelInfo";
            public const string publicationAdditionalInfoFolder = "AdditionalInformation";
            public const string publicationXLSFolder = "XLS";
            public const string publicationPhotoFolder = "Photos";
            public const string publicationReportsFolder = "Reports";
            public const string publicationFiguresFolder = "Figures";
            public const string publicationStylesFonts = "StyleFonts";

            //For legacy
            public const string mxdFolder_160415 = "MXD_Template";

        }

        public static class Styles
        {

            //Category names, used in style to find symbols
            public const string MapUnitCategory = "GSC-ArcGIS Shadeset";

            public const string DefaultLine_FGDC = "31.10"; //Will be used to get rid of <null> symbol values inside geoline
            public const string InvalidLine_FGDC = "19.01.06"; //Will be used to show invalid symbol selected by user, probably with manual edit of attribute table
            public const string InvalidPoint_FGDC = "18.67"; //Will be used to show invalid symbol selected by user, probably with manual edit of attribute table.
            public const string DefaultColorScheme = "Greens"; //A default value used to create UniqueValueRenderer, or else update method on symbols could crash of nothing is set.

            //Style files - Standard
            public const string DefaultStyleFileName = "GSC_SymbolStandard.style"; 

            //Embedded ressources
            public const string styleEmbeddedFolder = "Styles";
        }

        public static class Environment
        {

            #region New method using a database

            //Environment container name
            public const string containerName = "GSC_Project_Editor";
            public const string envRelPath = "\\ArcGIS\\";

            //Table names
            public const string envTable = "Environment";
            public const string qcNumData = "QC_NumericalDataReportTable";

            //Field names
            public const string envFolderPath = "PROJECT_WORKSPACE_PATH";
            public const string envCurrentDBPath = "PROJECT_DATABASE_PATH";
            public const string envCurrentDBName = "PROJECT_DATABASE_NAME";
            public const string envProjectScale = "PROJECT_SCALE";
            public const string envMapScale = "MAP_SCALE";
            public const string envLanguage = "LANGUAGE";
            public const string envProject = "PROJECT_NAME";
            public const string envActive = "IS_ACTIVE";
            public const string envReportType = "REPORT_TYPE";
            public const string envStyleMarkerField = "PROJECT_STYLE_MARKER_PATH";
            public const string envStyleLineField = "PROJECT_STYLE_LINE_PATH";
            public const string envStyleFillField = "PROJECT_STYLE_FILL_PATH";
            public const string envFieldControl = envFolderPath; ///UPDATE THIS VALUE IF WORKING ENV. DATABASE HAS CHANGED, THIS WILL BE USED AS A CONTROLLER TO DETECT NEW VERSIONS

            public const string qcNumData_DSName = "DATASET_NAME";
            public const string qcNumData_FName = "FIELD_NAME";
            public const string qcNumData_FType = "FIELD_TYPE";
            public const string qcNumData_FMin = "FIELD_MIN";
            public const string qcNumData_FMax = "FIELD_MAX";
            public const string qcNumData_FAvg = "FIELD_AVG";
            public const string qcNumData_FStrd = "FIELD_STD";
            public const string qcNumData_FCount = "FIELD_COUNT";

            #endregion

        }

        public static class Reports
        {
            //Path 
            public const string MetadataRelPath = "\\Style\\ReportMetadata.rlf";
        }

        public static class FieldDefaults
        {
            //Legend generator table
            public const string LegendSymbolType = "H2";

            //General
            public const string NotAvailable = "N.A.";

            //GeolineID
            public const string invalidGeolineIDValue = "000000000000";

            //GeopointID
            public const string invalidGeopointIDValue = "0000000000000";
        }

        public static class FieldValues
        {
            //Environment table field default
            public const int envActiveTrue = 1;
            public const int envActiveFalse = 0;
        }

        public static class ProjectOtherDatabase
        {
            public const string dbSuffix = "_RelatedData";
            public const string dbType = ".gdb";
        }

        public static class ProjectDatabaseType
        {
            public const string bedrockDB = "Bedrock";
            public const string surficialDB = "Surficial";
            public const string marineDB = "Marine";
        }

        public static class Fonts
        {
            #region BEDROCK
            public const string fgdcFont1 = "FGDCGeoSym01";
            public const string fgdcFont2 = "FGDCGeoSym02";
            public const string fgdcFont3 = "FGDCGeoSym03";
            public const string fgdcFont4 = "FGDCGeoSym04";
            public const string fgdcFont5 = "FGDCGeoSym05";

            public const string gscFont1 = "gsc1";
            public const string gscFont2 = "gsc2";
            public const string gscFont3 = "gsc3";
            public const string gscFont4 = "gsc4";
            public const string gscFont5 = "gsc5";

            public const string gscSymbolStandard1 = "GSCSymbolStandard1";

            public const string gscGeology = "gscGeology";
            #endregion

        }

        #endregion

        #region Other (Prog, math, etc.) related constants

        public static class ValueKeywords
        {
            public const string labelOverprint = "_OP";
            public const string GetUniqueFieldValuesMain = "Main";
            public const string FullProjectLegendSuffix = "Full_Project";
            //ESRI
            public const string esriInMemory = "GPInMemoryWorkspace";

            
        }

        public static class MathConstants
        {
            //Smallest area factor
            public const double smallAreaFactorClass1 = 1.5; //refers to 1.5mm on a paper map.
            public const double smallAreaFactorClass2 = 1;

        }

        public static class Debug
        {
            public const string debugFileName = "DebugLog.txt";
        }

        public static class Configuration
        {
            public const string userConfigFileName = "UserConfiguration.json";
        }

        public static class Seperator
        {
            public const string textFileLineSep = ";";
        }

        public static class Culture
        {
            public const string french = "fr";
            public const string english = "en";
        }

        /// <summary>
        ///Spatial relation description based on the 9 topologial relations
        ///https://en.wikipedia.org/wiki/DE-9IM
        ///http://resources.esri.com/help/9.3/arcgisengine/arcobjects/esriGeoDatabase/ISpatialFilter_SpatialRelDescription.htm
        /// </summary>
        public static class TopologicalRelations
        {
            public static string disjointQuery = "FF*FF****";
        }

        public static class ESRI
        {
            //http://resources.arcgis.com/en/help/arcobjects-net/conceptualhelp/index.html#//00010000029s000000

            public const string defaultArcGISFolderName = "ArcGIS";
        }
        #endregion

    }
}
