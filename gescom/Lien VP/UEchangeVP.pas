unit UEchangeVP;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ADODB, Hctrls,DB, UTOB,Hmsgbox,Hent1, UentCommun;

const categorie1='Capitalisation cadres';
      categorie2='Arrêt maladie';
      categorie3='6eme semaine production';
      categorie4='Congés payés légaux';
      categorie5='Congés d''ancienneté';
      categorie6='Congés exceptionnels';
      categorie7='Congés de maternité';
      categorie8='Congés de paternité';
      categorie9='Absences enfants malades';
      categorie10='Congé épargne temps (CET)';
      categorie11='RTT';
      categorie12='Arrêt accident du travail';
      categorie13='Absence crédit d''heures';
      categorie14='Absence sans solde';
      categorie15='Presence exceptionnelle';

      COMMUNE = 'BaseCommune';

type


  TEchangeVP = class (TObject)
  	private
      fServer : string;
      fDatabase : string;
      fEmplacementVp : string;
      //XX : TADOConnection;
      TheForm : TForm;
      //
    	function AddStringSql(Sql: string): string;
      function AddDateSql (laDate : TdateTime) : string;
    	function AddDateTimeSql(laDate: TdateTime): string;
    	function AddDoubleSql(valeur: Double; Plus : string = ''): string;
    	function BooleantoStr(Valeur: Boolean): string;
      //
    	procedure AjouteAbsences (TOBT : TOB);
      procedure AjouteHorairesRessource (Idressource : string;TOBT,TOBHrs : TOB;RessourcePhys : boolean);
      procedure AjouteHorairesSTD (TOBT : TOB);
      procedure AjouteMetier(TOBT : TOB);
    	procedure AjouteRessource(TOBT: TOB);
      procedure AjouteRessourceVirt (TOBT : TOB);
    	procedure AjouteTypeRessource(TOBT: TOB);
      procedure AssocieHoraire (TOBT : TOB;Code : string;CodeHrs : string; RessourcePhys : Boolean=true);
      //
    	function ConstitueHrs(TOBT: TOB): string;
      //
      //
    	function ExisteHorairesSTD(TOBT: TOB): Boolean;
      function ExisteMetier(TOBT : TOB): Boolean;
      function ExisteRessource(TOBT: TOB): Boolean; overload;
      function ExisteRessource(CodeRessource : string): Boolean; overload;
    	function ExisteTypeRessource(TOBT: TOB): Boolean;
      function ExisteHorairesRessource(IdRessource : string) : Boolean;
      function ExisteVPSql(Database : string;Sql: String): boolean;
      //
    	function GetOneresultSqlVP(Database : string; Sql: string): string;
    	function GetMultiResultSqlVP(Database,Sql: string): TOB;

      procedure MajRessource (TOBT : TOB);
      procedure MajHorairesSTD (TOBT : TOB);
      procedure MajHorairesRessource (Idressource : string ; TOBT,TOBHrs : TOB; RessourcePhys : boolean);
      procedure MajRessourceVirt (TOBT : TOB);
    	procedure SupprimeAbsences (DateDebut,DateFin : TdateTime);
    	procedure SetServeur(const Value: string);
      //
			procedure AjouteLienLSEVP(Libelle : string);
      function FindProjet (Libelle : string) : Integer;
      function GetServeur : string;
    	function constitueChaineConfigColonnesTaches: string;
      procedure EcritTache (XX : String ;TOBTT: TOB);
      function ConstitueSqlInsertTache (TOBTT : TOB) : string;
      function ConstitueAffectationtache (XX: string; Idtache : integer;TOBTT : TOB) : string;
			procedure MajDetailProjet (TOBTaches : TOB);
      procedure ConstitueLientache(XX : string;TOBTT,TOBTaches : TOB);
      procedure ConstitueRecette (XX : String;TOBTaches : TOB);
      procedure ConstitueDepense (XX : String;TOBTaches,TOBTT : TOB);
      procedure DeleteInfobase(CC : TADOCommand);
    	procedure SetForm(const Value: Tform);
    	function NormaliseStr(Chaine: string): string;
    	function IsliaisonPossible: Boolean;

      //
    public
      constructor create; overload;
      constructor create (TheServer : string); overload;
      destructor destroy; override;
      property Server : string read fServer write SetServeur;
      property Database : string read fDatabase write fDatabase;
      property ecran : Tform read TheForm write SetForm;
      property OkLiaison : Boolean read IsliaisonPossible;
      //
      //
      procedure AjouteProjet (TOBTaches : TOB);
    	procedure UpdateProjet(TOBTaches: TOB; WithSupress : Boolean=false);
      //
      function Existe : Boolean;
      function BaseExiste (IdBase : integer) : boolean;
      function GetEmplacementExe : string;
			procedure SetAbsencesSalaries (TOBAbsences : TOB; DateDebut,DateFin : TdateTime);
      procedure SetHorairesSTD (TOBHoraires : TOB);
      procedure SetInterv (TOBInterventions : TOB; DateDebut,DateFin : TdateTime);
			procedure SetMetiers (TOBMetiers : TOB);
      procedure SetRessources(TOBRessources: TOB);
    	procedure SetRessourcesVirt(Database: String; TOBPrestations: TOB);
    	procedure SetTypeRessource(TOBRessources: TOB);
      //
  end;

function ConstitueConnectionString (Serveur: string ; Database : string ) : string;
//function ConstitueConnection (ecran : TForm) : TADOConnection;
function GetEmplacementVPExe : string;

implementation
uses IniFiles,DialogEx,UtilTOBPiece,FactUtil,UdateUtils,FactVariante,
  	 DateUtils,Paramsoc,UTilFonctionCalcul,CPProcGen;

function GetEmplacementVPExe : string;
begin
  Result :=GetFromRegistry(HKEY_LOCAL_MACHINE,'Software\Microsoft\Windows\CurrentVersion\App Paths\VisualProjet4.exe','','',TRUE) ;
end;


{ TEchangeVP }

function ConstitueConnectionString(Serveur: string ; Database : string): string;
begin
  result := '';
  if Serveur = '' then
  begin
  	PGIInfo ('Veuillez définir le serveur HyperFile');
    exit;
  end;
	result := 'Provider=PCSoft.HFSQL';
  if Database <> '' then
  begin
  	Result := result + ';Initial Catalog='+DataBase;
  end;
  Result := result + ';User ID=Admin;Data Source='+Serveur+';Extended Properties="Language=ISO-8859-1"';
end;

constructor TEchangeVP.create;
begin
  Server := GetServeur;
end;

constructor TEchangeVP.create(TheServer: string);
begin
	Server := GetServeur;
  Server := TheServer;
end;

destructor TEchangeVP.destroy;
begin
  inherited;
end;

function TEchangeVP.BooleantoStr (Valeur : Boolean) : string;
begin
	if valeur then result := '1' else result := '0';
end;

function TEchangeVP.AddDateTimeSql(laDate: TdateTime): string;
begin
	Result := ''''+FormatDateTime('yyyymmddhhnnss', Ladate)+'''';
end;

function TEchangeVP.AddDateSql(laDate: TdateTime): string;
begin
	Result := ''''+FormatDateTime('yyyymmdd', Ladate)+'''';
end;

function TEchangeVP.AddStringSql(Sql : string) : string;
begin
	Result := ''''+StringReplace_(Sql,'''','''''',[rfReplaceAll])+'''';
end;

function TEchangeVP.AddDoubleSql (valeur : Double; Plus : string) : string;
begin
	Result := ''''+StringReplace_(STRFPOINT(VALEUR),'''','''''',[rfReplaceAll])+plus+'''';
end;


function TEchangeVP.Existe : Boolean;
begin
  Result := (GetEmplacementVPExe <> '');
end;

function TEchangeVP.ExisteVPSql(Database : string;Sql : String) : boolean;
var QQ : TADOQuery;
		ConnectionSt : string;
begin
  ConnectionSt := ConstitueConnectionString (fServer,Database);
 	QQ := TADOQuery.Create(TheForm);
  QQ.ConnectionString := ConnectionSt;
  QQ.LockType := ltReadOnly;
  TRY
    QQ.SQL.Add(SQL);
    QQ.Open;
		result := not QQ.eof;
    QQ.Active := false;
  finally
    QQ.Close;
    FreeAndNil(QQ);
  end;
end;

function TEchangeVP.ExisteMetier (TOBT : TOB) : Boolean;
var Sql : string;
    Metier : string;
begin
  Metier := AddStringSql(TOBT.GetString('BNP_LIBELLE'));
  //
  SQl := 'SELECT nomMetier FROM FichMetiers WHERE nomMetier='+ Metier;
  Result := ExisteVPSql(COMMUNE,Sql);
end;

function TechangeVp.NormaliseStr(Chaine : string) : string;
var stChaine : string;
begin
  StChaine := Trim(Chaine);
  ModifCarAccentueKeep(StChaine); // pour enlever les catacrères accentués
  SupprimeCaracteresSpeciaux(StChaine,false,True,True); // suppression des caractères spéciaux
	result:=UpperCase(StringReplace(StChaine,' ','_',[rfReplaceAll]));
end;

function TEchangeVP.ExisteRessource(CodeRessource : string): Boolean;
var Sql,Ressource : string;
begin
  Ressource := AddStringSql(NormaliseStr(Copy(CodeRessource,1,16)));
  //
  SQl := 'SELECT prenomNom FROM FichRessources WHERE trigramme='+ Ressource;
  Result := ExisteVPSql(COMMUNE,Sql);
end;

function TEchangeVP.ExisteRessource (TOBT : TOB) : Boolean;
var Code : string;
begin
  if TOBT.NomTable = 'ARTICLE' then
  begin
//  	Code := '_'+TOBT.GetString('GA_CODEARTICLE');
  	Code := NormaliseStr(Copy('_'+TOBT.GetString('GA_LIBELLE'),1,16));
  end else
  begin
//  	Code := TOBT.GetString('ARS_RESSOURCE');
		Code := NormaliseStr(Copy(TOBT.GetString('ARS_LIBELLE2')+' '+TOBT.GetString('ARS_LIBELLE'),1,16));
  end;
  Result := ExisteRessource(Code);
end;

function TEchangeVP.ExisteTypeRessource (TOBT : TOB) : Boolean;
var Sql : string;
    Metier : string;
begin
  Metier := AddStringSql(TOBT.GetString('CC_LIBELLE'));
  //
  SQl := 'SELECT nomMetier FROM FichMetiers WHERE nomMetier='+ Metier;
  Result := ExisteVPSql(COMMUNE,Sql);
end;

function TEchangeVP.ExisteHorairesRessource(IdRessource: string): Boolean;
var Sql : string;
begin
  //
  SQl := 'SELECT IDPlageHoraire FROM FichPlagesHoraires WHERE IDressource='+ AddStringSql(IdRessource);
  Result := ExisteVPSql(COMMUNE,Sql);
end;

function TEchangeVP.ExisteHorairesSTD (TOBT : TOB) : Boolean;
var Sql : string;
    HoraireStd : string;
begin
  HoraireStd := AddStringSql('STD');
  //
  SQl := 'SELECT nomPlageHoraire FROM FichModelePlagesHoraires WHERE nomPlageHoraire='+ HoraireStd;
  Result := ExisteVPSql(COMMUNE,Sql);
end;

function TEchangeVP.ConstitueHrs (TOBT : TOB) : string;
begin
  Result := '1,'+BooleantoStr(not GetParamSocSecur('SO_JOURFERMETURE1',False))+',1.00,'+FormatDateTime('hh:mm',TOBT.GetValue('DEBAM'))+'-'+
            FormatDateTime('hh:mm',TOBT.GetValue('FINAM'))+','+
            FormatDateTime('hh:mm',TOBT.GetValue('DEBUTPM'))+'-'+
            FormatDateTime('hh:mm',TOBT.GetValue('FINPM'))+#13#10 +
  					'2,'+BooleantoStr(not GetParamSocSecur('SO_JOURFERMETURE2',False))+',1.00,'+FormatDateTime('hh:mm',TOBT.GetValue('DEBAM'))+'-'+
            FormatDateTime('hh:mm',TOBT.GetValue('FINAM'))+','+
            FormatDateTime('hh:mm',TOBT.GetValue('DEBUTPM'))+'-'+
            FormatDateTime('hh:mm',TOBT.GetValue('FINPM'))+#13#10 +
            '3,'+BooleantoStr(not GetParamSocSecur('SO_JOURFERMETURE3',False))+',1.00,'+FormatDateTime('hh:mm',TOBT.GetValue('DEBAM'))+'-'+
            FormatDateTime('hh:mm',TOBT.GetValue('FINAM'))+','+
            FormatDateTime('hh:mm',TOBT.GetValue('DEBUTPM'))+'-'+
            FormatDateTime('hh:mm',TOBT.GetValue('FINPM'))+#13#10 +
            '4,'+BooleantoStr(not GetParamSocSecur('SO_JOURFERMETURE4',False))+',1.00,'+FormatDateTime('hh:mm',TOBT.GetValue('DEBAM'))+'-'+
            FormatDateTime('hh:mm',TOBT.GetValue('FINAM'))+','+
            FormatDateTime('hh:mm',TOBT.GetValue('DEBUTPM'))+'-'+
            FormatDateTime('hh:mm',TOBT.GetValue('FINPM'))+#13#10 +
            '5,'+BooleantoStr(not GetParamSocSecur('SO_JOURFERMETURE5',False))+',1.00,'+FormatDateTime('hh:mm',TOBT.GetValue('DEBAM'))+'-'+
            FormatDateTime('hh:mm',TOBT.GetValue('FINAM'))+','+
            FormatDateTime('hh:mm',TOBT.GetValue('DEBUTPM'))+'-'+
            FormatDateTime('hh:mm',TOBT.GetValue('FINPM'))+#13#10+
            '6,'+BooleantoStr(not GetParamSocSecur('SO_JOURFERMETURE6',False))+',1.50,'+FormatDateTime('hh:mm',TOBT.GetValue('DEBAM'))+'-'+
            FormatDateTime('hh:mm',TOBT.GetValue('FINAM'))+','+
            FormatDateTime('hh:mm',TOBT.GetValue('DEBUTPM'))+'-'+
            FormatDateTime('hh:mm',TOBT.GetValue('FINPM'))+#13#10 +
            '7,'+BooleantoStr(not GetParamSocSecur('SO_JOURFERMETURE',False))+',1.50,'+FormatDateTime('hh:mm',TOBT.GetValue('DEBAM'))+'-'+
            FormatDateTime('hh:mm',TOBT.GetValue('FINAM'))+','+
            FormatDateTime('hh:mm',TOBT.GetValue('DEBUTPM'))+'-'+
            FormatDateTime('hh:mm',TOBT.GetValue('FINPM'));
end;

procedure TEchangeVP.AjouteAbsences (TOBT : TOB);
var TT : TADOCommand;
    Sql,Code,IdRessource,NomAbsence : string;
    heureD,heureF : TDateTime;
    ConnectionSt : string;
begin
  ConnectionSt := ConstitueConnectionString (fServer,COMMUNE);
  if TOBT.NomTable = 'ABSENCESALARIE' then
  begin
  	Code := TOBT.GetString('ARS_RESSOURCE');
  end else
  begin
  	Code := TOBT.GetString('BEP_RESSOURCE');
  end;
  Code := NormaliseStr(Copy(Code,1,16));
  If not ExisteRessource(Code) then Exit; // au cas ou
  Sql := 'SELECT IDressource FROM FichRessources WHERE trigramme='+AddStringSql(Code);
  Idressource := GetOneresultSqlVP(COMMUNE,Sql);
  //
  if TOBT.NomTable = 'ABSENCESALARIE' then
  begin
    if TobT.GetString('PCN_TYPECONGE') = 'PRI' then
      NomAbsence := categorie4
    else if TobT.GetString('PCN_TYPECONGE') = 'CPA' then
      NomAbsence := categorie4
    else if TobT.GetString('PCN_TYPECONGE') = 'INT' then
      NomAbsence := 'Absence Intempérie'
    else if TobT.GetString('PCN_TYPECONGE') = 'MAL' then
      NomAbsence := categorie2
    else if TobT.GetString('PCN_TYPECONGE') = 'MAT' then
      NomAbsence := categorie7
    else if TobT.GetString('PCN_TYPECONGE') = 'PAT' then
      NomAbsence := categorie8
    else if TobT.GetString('PCN_TYPECONGE') = 'RTT' then
      NomAbsence := categorie11
    else
      NomAbsence := 'Absences diverses';
  end else
  begin
      NomAbsence := 'Intervention';
  end;
  //

  if TOBT.NomTable = 'ABSENCESALARIE' then
  begin
    // -------------------------------------------------------
    // positionnement des heures de debut et fin d'absence
    // -------------------------------------------------------
    if TobT.getValue('PCN_DEBUTDJ')='MAT' then
    begin
      heureD := GetDebutMatinee;
    end else
    begin
      heureD := GetDebutApresMidi;
    end;
    TOBT.SetDateTime('PCN_DATEDEBUTABS',StrToDate(DateToStr(TOBT.GetDateTime('PCN_DATEDEBUTABS')))+HeureD);
    if TobT.getValue('PCN_FINDJ')='MAT' then
    begin
      heureF := GetFinMatinee;
    end else
    begin
      heureF := GetFinApresMidi;
    end;
    TOBT.SetDateTime('PCN_DATEFINABS',StrToDate(DateToStr(TOBT.GetDateTime('PCN_DATEFINABS')))+HeureF);
    // -------------------------------------------------------
    Sql := 'INSERT INTO FichAbsences (nomAbsence,IDressource,dateHeureDeb,dateHeureFin,estPresent) values '+
          '('+AddStringSql(nomAbsence)+','+
            AddStringSql(Idressource)+','+
            AddDateTimeSql(TobT.GetDateTime ('PCN_DATEDEBUTABS'))+','+
            AddDateTimeSql(TobT.GetDateTime('PCN_DATEFINABS'))+',0)';
  end else
  begin
    Sql := 'INSERT INTO FichAbsences (nomAbsence,IDressource,dateHeureDeb,dateHeureFin,estPresent) values '+
          '('+AddStringSql(nomAbsence)+','+
            AddStringSql(Idressource)+','+
            AddDateTimeSql(TobT.GetDateTime ('BEP_DATEDEB'))+','+
            AddDateTimeSql(TobT.GetDateTime('BEP_DATEFIN'))+',0)';
  end;
  //
  TT := TADOCommand.Create(TheForm);
  try
//    TT.Connection := XX;
		TT.ConnectionString := ConnectionSt;
    TT.CommandText := SQL;
    TT.Execute;
  finally
    TT.Free;
  end;

end;


procedure TEchangeVP.AjouteHorairesRessource(Idressource: string; TOBT, TOBHrs: TOB; RessourcePhys : boolean);
var TT : TADOCommand;
    Sql,ChaineHrs,Valeur,ConnectionSt : string;
begin
  ConnectionSt := ConstitueConnectionString (fServer,COMMUNE);
  ChaineHrs := AddStringSql(TOBHrs.GetString('chainePlagesHoraires'));
  if not RessourcePhys then
  begin
  	valeur := AddDoubleSql(TOBT.GetValue('GA_DPR'));
  end else
  begin
  	valeur := AddDoubleSql(TOBT.GetValue('ARS_TAUXREVIENTUN'));
  end;
  //
  Sql := 'INSERT INTO FichPlagesHoraires (IDressource,dateDebPeriode,dateFinPeriode,coutHoraire,chainePlagesHoraires) values '+
  			'('+AddStringSql(IdRessource)+','+AddDateSql(StrToDate('01/01/1900'))+','+AddDateSql(StrToDate('31/12/2999'))+','+Valeur+','+ChaineHrs+')';
  TT := TADOCommand.Create(TheForm);
  try
    TT.ConnectionString := ConnectionSt;
    TT.CommandText := SQL;
    TT.Execute;
  finally
    FreeAndNil(TT);
  end;
end;

procedure TEchangeVP.AjouteHorairesSTD(TOBT: TOB);
var TT : TADOCommand;
    Sql,HoraireStd,ChaineVide,ChaineHrs,ConnectionSt : string;
begin
  ConnectionSt := ConstitueConnectionString (fServer,COMMUNE);
  //
  HoraireStd := AddStringSql('STD');
  ChaineVide := AddStringSql('');
  ChaineHrs := AddStringSql(ConstitueHrs(TOBT));
  //
  Sql := 'INSERT INTO FichModelePlagesHoraires (NomPlageHoraire,detailsPlageHoraire,chainePlagesHoraires) values '+
  			'('+HoraireStd+','+ChaineVide+','+ChaineHrs+')';
  TT := TADOCommand.Create(TheForm);
  try
    TT.ConnectionString  := ConnectionSt;
    TT.CommandText := SQL;
    TT.Execute;
  finally
    FreeAndNil(TT);
  end;
end;

procedure TEchangeVP.AjouteMetier (TOBT : TOB);
var TT : TADOCommand;
    Sql,Metier : string;
    ConnectionSt : string;
begin
  ConnectionSt := ConstitueConnectionString (fServer,COMMUNE);
  Metier :=AddStringSql(TOBT.GetString('BNP_LIBELLE'));

  Sql := 'INSERT INTO FichMetiers (nomMetier,detailsMetier) values ('+Metier+','+Metier+')';
  TT := TADOCommand.Create(TheForm);
  try
    TT.ConnectionString  := ConnectionSt;
    TT.CommandText := SQL;
    TT.Execute;
  finally
    FreeAndNil(TT);
  end;
end;

procedure TEchangeVP.AjouteTypeRessource (TOBT : TOB);
var TT : TADOCommand;
    Sql,Metier,ConnectionSt : string;
begin
  ConnectionSt := ConstitueConnectionString (fServer,COMMUNE);
  Metier :=AddStringSql(TOBT.GetString('CC_LIBELLE'));
  Sql := 'INSERT INTO FichMetiers (nomMetier,detailsMetier) values ('+Metier+','+Metier+')';
  TT := TADOCommand.Create(TheForm);
  try
    TT.ConnectionString := ConnectionSt;
    TT.CommandText := SQL;
    TT.Execute;
  finally
    FreeAndNil(TT);
  end;
end;

procedure TEchangeVP.AssocieHoraire(TOBT: TOB; Code, CodeHrs: string; RessourcePhys : Boolean=true);
var TOBHrs : TOB;
		Sql : string;
    Idressource : string;
begin
  Sql := 'SELECT IDressource FROM FichRessources WHERE trigramme='+Code;
  Idressource := GetOneresultSqlVP(COMMUNE,Sql);
  //
	Sql := 'SELECT IDModelePlageHoraire,chainePlagesHoraires FROM FichModelePlagesHoraires '+
  			 'WHERE nomPlageHoraire='+AddStringSql(CodeHrs);
	TOBHrs := GetMultiResultSqlVP (COMMUNE,Sql);
  TRY
    if not ExisteHorairesRessource(IdRessource) then
    begin
      AjouteHorairesRessource (Idressource,TOBT,TOBHrs,RessourcePhys);
    end else
    begin
      MajHorairesRessource (Idressource,TOBT,TOBHrs,RessourcePhys);
    end;

  FINALLY
    TOBHrs.free;
  END;
end;

procedure TEchangeVP.AjouteRessourceVirt (TOBT : TOB);
var TT : TADOCommand;
    ConnectionSt,Sql,Ressource,Code,CodePrest,IdMetier,ChaineVide,Metier,EstMaterielle : string;
begin
  ConnectionSt := ConstitueConnectionString (fServer,COMMUNE);
  ChaineVide := '''''';
  if TOBT.GetString('GA_LIBELLE') = '' then exit;

	Metier := AddStringSql(Trim(Copy(TOBT.GetValue('BNP_LIBELLE'),1,50)));
  Ressource := AddStringSql(NormaliseStr(Copy('_'+TOBT.GetString('GA_LIBELLE'),1,16)));
  //
//  CodePrest := Copy('_'+TOBT.GetString('GA_CODEARTICLE'),1,16);
  CodePrest := NormaliseStr(Copy('_'+TOBT.GetValue('GA_LIBELLE'),1,16));
  Code := AddStringSql(CodePrest);
  if TOBT.GetString('BNP_TYPERESSOURCE')='MAT' then EstMaterielle :='1'
  																						 else EstMaterielle :='0';
  //
  Sql := 'SELECT IdMetier FROM FichMetiers WHERE nomMetier='+ Metier;
  IdMetier := GetOneresultSqlVP (COMMUNE,Sql); if IdMetier = '' then IdMetier := ChaineVide;
  //
  Sql := 'INSERT INTO FichRessources '+
  				'(prenomNom,trigramme,password,IdEquipe,IDSite,IDMEtier,DetailsRessource,'+
          'nomFichPhoto,numTelBureau,numTelMobile,mail,estInactive,estMaterielle,chaineNumFoncBloquees) '+
          'values ('+
          Ressource+','+
          Code+','+
          ChaineVide+',-1,-1,'+
          IdMetier+','+
          ChaineVide+','+
          ChaineVide+','+
          ChaineVide+','+
          ChaineVide+','+
          ChaineVide+
          ',0,'+
          EstMaterielle+','+
          ChaineVide+')';
  TT := TADOCommand.Create(TheForm);
  try
    TT.ConnectionString := ConnectionSt;
    TT.CommandText := SQL;
    TT.Execute;
  finally
    FreeAndNil(TT);
    AssocieHoraire (TOBT,Code,'STD',false);
  end;
end;

procedure TEchangeVP.AjouteRessource (TOBT : TOB);
var TT : TADOCommand;
    ConnectionSt,Sql,Ressource,Code,IdMetier,ChaineVide,Metier,EstMaterielle : string;
begin
  ConnectionSt := ConstitueConnectionString (fServer,COMMUNE);
  (*
  if TOBT.GetString('BNP_LIBELLE2')='' then
  begin
  	Metier := AddStringSql(TOBT.GetString('BNP_LIBELLE'));
  end else
  begin
  	Metier := AddStringSql(TOBT.GetString('BNP_LIBELLE2'));
  end;
  *)
  Metier := AddStringSql(TOBT.GetString('METIER'));
  ChaineVide := '''''';
  //
  if TOBT.GetString('ARS_TYPERESSOURCE')='MAT' then EstMaterielle :='1'
  																						 else EstMaterielle :='0';
  Ressource := AddStringSql(TOBT.GetString('ARS_LIBELLE2')+' '+TOBT.GetString('ARS_LIBELLE'));
//  Code := AddStringSql(Copy(TOBT.GetString('ARS_RESSOURCE'),1,16));
	Code := AddStringSql(NormaliseStr(Copy(TOBT.GetString('ARS_LIBELLE2')+' '+TOBT.GetString('ARS_LIBELLE'),1,16)));
  Sql := 'SELECT IdMetier FROM FichMetiers WHERE nomMetier='+ Metier;
  IdMetier := GetOneresultSqlVP (COMMUNE,Sql); if IdMetier = '' then IdMetier := ChaineVide;
  Sql := 'INSERT INTO FichRessources '+
  				'(prenomNom,trigramme,password,IdEquipe,IDSite,IDMEtier,DetailsRessource,'+
          'nomFichPhoto,numTelBureau,numTelMobile,mail,estInactive,estMaterielle,chaineNumFoncBloquees) '+
          'values ('+Ressource+','+Code+','+ChaineVide+',-1,-1,'+IdMetier+','+
          ChaineVide+','+ChaineVide+','+ChaineVide+','+ChaineVide+','+ChaineVide+
          ',0,'+EstMaterielle+','+ChaineVide+')';
  TT := TADOCommand.Create(TheForm);
  try
    TT.ConnectionString  := ConnectionSt;
    TT.CommandText := SQL;
    TT.Execute;
  finally
    FreeAndNil(TT);
    AssocieHoraire (TOBT,Code,'STD');
  end;
end;
(*
function ConstitueConnection (ecran : TForm) : TadoConnection;
begin
  result := TADOConnection.Create(ecran);
  result.ConnectOptions := coAsyncConnect;
  result.KeepConnection := true;
  result.LoginPrompt := False;
  result.Mode := cmReadWrite;
end;
*)
function TEchangeVP.GetMultiResultSqlVP(Database,Sql: string): TOB;
var OneQuery : TADOQuery;
    i : Integer;
    ConnectionSt : string;
begin
  ConnectionSt := ConstitueConnectionString (fServer,Database);
  result := TOB.Create ('LE RESULT',nil,-1);
 	OneQuery := TADOQuery.Create(TheForm);
    OneQuery.ConnectionString  := ConnectionSt;
  	OneQuery.LockType := ltReadOnly;
    OneQuery.SQL.Add(SQL);
  TRY
    OneQuery.Open;
		if not OneQuery.eof then
    begin
      for I := 0 to OneQuery.FieldCount -1 do
      begin
        result.AddChampSupValeur(OneQuery.FieldList[i].fieldname ,OneQuery.FieldList[i].Value);
      end;
    end;
  finally
  	OneQuery.Active := false;
    OneQuery.Close;
    FreeAndNil(OneQuery);
  end;
end;

function TEchangeVP.GetOneresultSqlVP (Database : string; Sql : string) : string;
var OneQuery : TADOQuery;
		ConnectionSt : string;
begin
  Result := '';
  ConnectionSt := ConstitueConnectionString (fServer,Database);
 	OneQuery := TADOQuery.Create(TheForm);
  TRY
    OneQuery.ConnectionString  := ConnectionSt;
  	OneQuery.LockType := ltReadOnly;
    OneQuery.SQL.Add(SQL);
    OneQuery.Open;
		if not OneQuery.eof then
    begin
			Result := OneQuery.FieldList[0].AsString;
    end;
  finally
  	OneQuery.Active := false;
    OneQuery.Close;
    FreeAndNil(OneQuery);
  end;
end;


procedure TEchangeVP.SupprimeAbsences (DateDebut,DateFin : TdateTime);

  function ConstitueRequeteAbsence (DateD,DateF : TdateTime) : string ;
  begin
    result := '( '+ // 1
              '(DateHeureDeb>=' + AddDateTimeSql(DateD) + ' AND DateHeureDeb<=' + AddDateTimeSql(DateF) + ') OR '+
              '(DateheureFin>=' + AddDateTimeSql(DateD) + ' AND DateheureFin<=' + AddDateTimeSql(DateF) + ') OR ' +
              '(DateheureDeb<=' + AddDateTimeSql(DateD) + ' AND DateheureFin>=' + AddDateTimeSql(DateF) + ')' +
              ')'; //1
  end;

var TT : TADOCommand;
    ConnectionSt,Sql : string;
begin
  ConnectionSt := ConstitueConnectionString (fServer,Database);
  Sql := 'DELETE FROM FichAbsences WHERE '+ConstitueRequeteAbsence (DateDebut,DateFin);
  TT := TADOCommand.Create(TheForm);
  try
    TT.ConnectionString  := ConnectionSt;
    TT.CommandText := SQL;
    TT.Execute;
  finally
    FreeAndNil(TT);
  end;
end;


procedure TEchangeVP.SetAbsencesSalaries(TOBAbsences: TOB; DateDebut,DateFin : TdateTime);
var i : Integer;
    TOBT : TOB;
begin
  SupprimeAbsences (DateDebut,DateFin);

  for I := 0 to TOBAbsences.detail.count -1 do
  begin
    TOBT := TOBAbsences.detail[I];
    AjouteAbsences (TOBT);
  end;
end;

procedure TEchangeVP.SetHorairesSTD(TOBHoraires: TOB);
var i : Integer;
    TOBT : TOB;
begin
  for I := 0 to TOBHoraires.detail.count -1 do
  begin
    TOBT := TOBHoraires.detail[I];
    if not ExisteHorairesSTD(TOBT) then
    begin
      AjouteHorairesSTD (TOBT);
    end else
    begin
      MajHorairesSTD (TOBT);
    end;
  end;
end;

procedure TEchangeVP.SetInterv (TOBInterventions : TOB; DateDebut,DateFin : TdateTime);
var i : Integer;
    TOBT : TOB;
begin
  for I := 0 to TOBInterventions.detail.count -1 do
  begin
    TOBT := TOBInterventions.detail[I];
    AjouteAbsences (TOBT);
  end;
end;

procedure TEchangeVP.SetMetiers(TOBMetiers: TOB);
var i : Integer;
    TOBT : TOB;
begin
  for I := 0 to TOBMetiers.detail.count -1 do
  begin
    TOBT := TOBMetiers.detail[I];
    if not ExisteMetier(TOBT) then
    begin
      AjouteMetier (TOBT);
    end;
  end;
end;

procedure TEchangeVP.SetRessources(TOBRessources: TOB);
var i : Integer;
    TOBT : TOB;
    {$IFDEF DEMOVP}
    Metier : string;
    Nbr : Integer;
    {$ENDIF}
begin
  //
{$IFDEF DEMOVP}
  metier := '';
  Nbr := 1;
{$ENDIF}
  for I := 0 to TOBRessources.detail.count -1 do
  begin
    TOBT := TOBRessources.detail[I];
{$IFDEF DEMOVP}
    if Metier <> TOBT.GetString('METIER') then
    begin
      Nbr :=1;
      Metier := TOBT.GetString('METIER');
    end else
    begin
      Inc(Nbr);
      if Nbr > 10 then continue;
    end;
{$ENDIF}
    if not ExisteRessource(TOBT) then
    begin
      AjouteRessource (TOBT);
    end else
    begin
      MajRessource(TOBT);
    end;
  end;
end;

procedure TEchangeVP.SetRessourcesVirt(Database: string; TOBPrestations: TOB);
var i : Integer;
    TOBT : TOB;
    CodeRessource : string;
begin
  for I := 0 to TOBPrestations.detail.count -1 do
  begin
    TOBT := TOBPrestations.detail[I];
  	CodeRessource := Copy('_'+TOBT.GetString('GA_CODEARTICLE'),1,50);
    if not ExisteRessource(CodeRessource) then
    begin
      AjouteRessourceVirt (TOBT);
    end else
    begin
      MajRessourceVirt (TOBT);
    end;
  end;
end;

procedure TEchangeVP.SetTypeRessource(TOBRessources: TOB);
var i : Integer;
    TOBT : TOB;
begin
  for I := 0 to TOBRessources.detail.count -1 do
  begin
    TOBT := TOBRessources.detail[I];
    if not ExisteTypeRessource(TOBT) then
    begin
      AjouteTypeRessource (TOBT);
    end;
  end;
end;

procedure TEchangeVP.MajRessource (TOBT : TOB);
var TT : TADOCommand;
    ConnectionSt,Sql,Ressource,Code,IdMetier,ChaineVide,Metier : string;
begin
//
  ConnectionSt := ConstitueConnectionString (fServer,COMMUNE);
  ChaineVide := '''''';
//
	(*
  if TOBT.GetString('BNP_LIBELLE2')='' then
  begin
  	Metier := AddStringSql(TOBT.GetString('BNP_LIBELLE'));
  end else
  begin
  	Metier := AddStringSql(TOBT.GetString('BNP_LIBELLE2'));
  end;
  *)
  Metier := AddStringSql(TOBT.GetString('METIER'));
	Code := AddStringSql(NormaliseStr(Copy(TOBT.GetString('ARS_LIBELLE2')+' '+TOBT.GetString('ARS_LIBELLE'),1,16)));
  Sql := 'SELECT IdMetier FROM FichMetiers WHERE nomMetier='+ Metier;
  IdMetier := GetOneresultSqlVP (COMMUNE,Sql); if IdMetier = '' then IdMetier := ChaineVide;
  Sql := 'UPDATE FichRessources SET IdMetier='+IdMetier+' WHERE trigramme='+Code ;
  TT := TADOCommand.Create(TheForm);
  try
    TT.ConnectionString  := ConnectionSt;
    TT.CommandText := SQL;
    TT.Execute;
  finally
    FreeAndNil(TT);
    AssocieHoraire (TOBT,Code,'STD');
  end;
end;

procedure TEchangeVP.MajRessourceVirt (TOBT : TOB);
var Ressource,Metier,IdMetier,SQL,ChaineVide,ConnectionSt : string;
		TT : TADOCommand;
begin
  ConnectionSt := ConstitueConnectionString (fServer,COMMUNE);
  ChaineVide := '''''';
  ressource := AddStringSql(NormaliseStr(Copy('_'+TOBT.GetString('GA_LIBELLE'),1,16)));
	Metier := AddStringSql(Trim(Copy(TOBT.GetValue('BNP_LIBELLE'),1,50)));
  Sql := 'SELECT IdMetier FROM FichMetiers WHERE nomMetier='+ Metier;
  IdMetier := GetOneresultSqlVP (COMMUNE,Sql); if IdMetier = '' then IdMetier := ChaineVide;
  Sql := 'UPDATE FichRessources SET IdMetier='+IdMetier;
  TT := TADOCommand.Create(TheForm);
  try
    TT.ConnectionString := ConnectionSt;
    TT.CommandText := SQL;
    TT.Execute;
  finally
    FreeAndNil(TT);
	  AssocieHoraire (TOBT,ressource,'STD',false);
  end;
end;


procedure TEchangeVP.MajHorairesSTD(TOBT: TOB);
var TT : TADOCommand;
    ConnectionSt,Sql,HoraireStd,ChaineVide,ChaineHrs : string;
begin
  ConnectionSt := ConstitueConnectionString (fServer,COMMUNE);
  HoraireStd := AddStringSql('STD');
  ChaineVide := AddStringSql('');
  ChaineHrs := AddStringSql(ConstitueHrs(TOBT));
  //
  Sql := 'UPDATE FichModelePlagesHoraires SET chainePlagesHoraires='+ChaineHrs+' WHERE nomPlageHoraire='+HoraireStd ;
  TT := TADOCommand.Create(TheForm);
  try
    TT.ConnectionString := ConnectionSt;
    TT.CommandText := SQL;
    TT.Execute;
  finally
    FreeAndNil(TT);
  end;
end;

procedure TEchangeVP.MajHorairesRessource(Idressource: string; TOBT, TOBHrs: TOB; RessourcePhys : boolean);
var TT : TADOCommand;
    ConnectionSt,Sql,ChaineHrs,Valeur : string;
begin
  ConnectionSt := ConstitueConnectionString (fServer,COMMUNE);
  ChaineHrs := AddStringSql(TOBHrs.GetString('chainePlagesHoraires'));
  if not RessourcePhys then
  begin
  	valeur := AddDoubleSql(TOBT.GetValue('GA_DPR'));
  end else
  begin
  	valeur := AddDoubleSql(TOBT.GetValue('ARS_TAUXREVIENTUN'));
  end;
  //
  Sql := 'UPDATE FichPlagesHoraires SET CoutHoraire='+valeur+',chainePlagesHoraires='+ChaineHrs+' WHERE '+
         'IDressource='+AddStringSql(IdRessource);
  TT := TADOCommand.Create(TheForm);
  try
    TT.ConnectionString  := ConnectionSt;
    TT.CommandText := SQL;
    TT.Execute;
  finally
    FreeAndNil(TT);
  end;
end;

function TEchangeVP.GetEmplacementExe : string;
begin
  Result :=GetEmplacementVPExe;
end;


procedure TEchangeVP.SetServeur(const Value: string);
begin
  fServer := Value;
end;

procedure TEchangeVP.AjouteProjet(TOBTaches : TOB);
var IdProjet : Integer;
begin
  AjouteLienLSEVP(TOBTaches.GetString('REFERENCE'));
  IdProjet := FindProjet (TOBTaches.GetString('REFERENCE'));
  if IdProjet <> -1 then
  begin
    TobTaches.SetInteger('IDPROJET',IdProjet);
    Updateprojet (TOBTaches);
  end;
end;

procedure TEchangeVP.AjouteLienLSEVP(Libelle: string);
var CurDir,VpDir,ExeVp : string;
begin
  Curdir := ExtractFileDir(Application.exename);
	VpDir := ExtractFileDir(GetEmplacementExe);
  if VpDir = '' then Exit;
  ExeVp := IncludeTrailingBackslash(VpDir)+'LSEcreeProjetVP.exe';
  if not FileExists(exeVp) then Exit;
  ChDir(VPDIR);
	FileExecAndWait (EXEVP+' "'+Libelle+'"');
  ChDir(CurDir);
end;

function TEchangeVP.FindProjet(Libelle: string): Integer;
var OneQuery : TADOQuery;
		ConnectionSt : string;
    Sql : string;
begin
  Result := -1;
 	OneQuery := TADOQuery.Create(TheForm);
  ConnectionSt := ConstitueConnectionString (fServer,COMMUNE);
  OneQuery.LockType := ltReadOnly;
  OneQuery.ConnectionString  := ConnectionSt;
  SQL := 'SELECT Idprojet FROM FichListeProjets WHERE NomProjet='+AddStringSql(Libelle);
  OneQuery.SQL.Add(SQL);
  TRY
    OneQuery.Open;
		if not OneQuery.eof then
    begin
      Result := OneQuery.findField('Idprojet').AsInteger;
    end;
  finally
  	OneQuery.Active := false;
    OneQuery.Close;
    FreeAndNil(OneQuery);
  end;
end;

function TEchangeVP.constitueChaineConfigColonnesTaches : string;
begin
  result := '1,IDtache,ID,0,17,0'+#13#10+
  					'2,numTache,n°,1,22,1'+#13#10+
            '3,niveauTache,Niveau,0,40,0'+#13#10+
            '4,pictosTache,pictos,0,34,0'+#13#10+
            '5,etatRecapReplie,+/-,1,20,1'+#13#10+
            '6,nomTache,Nom de la tâche,1,174,1'+#13#10+
            '7,travailAttente,Travail/Attente,1,77,1'+#13#10+
            '8,Complements,Compléments,0,69,0'+#13#10+
            '9,totalTravail,Travail re-prévu,0,80,0'+#13#10+
            '10,totalAttente,Attente re-prévue,0,88,0'+#13#10+
            '11,dureeTache,Durée,1,42,1'+#13#10+
            '12,realise,Réalisé,0,41,0'+#13#10+
            '13,ressources,Ressources,1,78,1'+#13#10+
            '14,tachesPrec,Taches prec,0,66,0'+#13#10+
            '15,tachesSuiv,Taches suiv,0,64,0'+#13#10+
            '16,dateHeureDeb,Début,1,90,1'+#13#10+
            '17,dateHeureFin,Fin,1,90,1'+#13#10+
            '18,dateHeureDebMini,Début pas avant,0,90,0'+#13#10+
            '19,dateHeureFinMaxi,Fin pas après,1,90,1'+#13#10+
            '20,retardAvance,Retard/avance,0,79,0'+#13#10+
            '21,totalMargeLibre,Marge libre,0,58,0'+#13#10+
            '22,IDCategorieTache,Code catégorie,0,78,0'+#13#10+
            '23,dateHeureDebReel,Commencée le,0,76,0'+#13#10+
            '24,dateHeureFinReelle,Terminée le,0,61,0'+#13#10+
            '25,tauxAchevement,Avancement,1,67,1'+#13#10+
            '26,totalCout,Total coût horaire,0,89,0'+#13#10+
            '27,detailsTache,Détails tâche,0,68,0'+#13#10+
            '28,resteAfaire,Reste à faire,1,80,0';
end;

procedure TEchangeVP.UpdateProjet(TOBTaches : TOB; WithSupress : Boolean=false);

	procedure  AjouteRessource(TOBTT,TOBressProjet : TOB);
  var TOBRP : TOB;
  		Coderessource : string;
  begin
//  	Coderessource := Copy('_'+TOBTT.GetString('GL_CODEARTICLE'),1,16);
		if TOBTT.GetString('GA_LIBELLE')='' then Exit;
    AjouteRessourceVirt (TOBTT);
  	Coderessource := NormaliseStr(Copy('_'+TOBTT.GetString('GA_LIBELLE'),1,16));
		TOBRP := TOB.Create('UNE AFFECT',TOBressProjet,-1);
    TOBRP.AddChampSupValeur('RESSOURCE',Coderessource);
  end;

	function IsExisteRess (TOBTT,TOBressProjet: TOB ) : Boolean;
  var Coderessource : string;
  begin
//  	Coderessource := Copy('_'+TOBTT.GetString('GL_CODEARTICLE'),1,16);
  	Coderessource := NormaliseStr(Copy('_'+TOBTT.GetString('GA_LIBELLE'),1,16));
    result := (TOBressProjet.FindFirst(['RESSOURCE'],[Coderessource],true) <> nil);
  end;

	procedure ConstitueressourceProjet (Database : string; TOBTaches,TOBressProjet : TOB);
  var II : Integer;
  		TOBTT : TOB;
  begin
  	For II := 0 to TOBTaches.detail.count -1 do
    begin
			TOBTT := TOBTaches.detail[II];
			if TOBTT.GetString('GA_LIBELLE')='' then Continue;
      if IsVariante(TOBTT) then Continue;
      if IsExisteRess (TOBTT,TOBressProjet) then
      begin
      	MajRessourceVirt (TOBTT);
      end else
      begin
      	AjouteRessource(TOBTT,TOBressProjet);
      end;
    end;
  end;

  procedure AffecteRessourceprojet(IdProjet : integer; TOBressProjet : TOB );
  var QQ : TAdoQuery;
  		TT : TADOCommand;
  		ii : Integer;
      ConnectionST,idRess,Sql : string;
      TOBRP : TOB;
  begin
    // nettoyage
  	ConnectionSt := ConstitueConnectionString (fServer,COMMUNE);
    TT := TADOCommand.Create(nil);
    TT.ConnectionString := ConnectionSt;
    TRY
    	TT.CommandText := 'DELETE FROM FichRessProjets WHERE IdProjet='+inttoStr(IdProjet);
  		TT.Execute;
    finally
      FreeAndNil(TT);
    end;
    //
    TRY
      for ii := 0 to TOBressProjet.detail.Count -1 do
      begin
				TOBRP :=  TOBressProjet.detail[II];
        QQ := TADOQuery.Create(TheForm);
        QQ.ConnectionString := ConnectionST;
        QQ.LockType := ltReadOnly;
        QQ.SQL.Clear;
        QQ.SQL.Add('SELECT Idressource FROM FichRessources WHERE trigramme='+AddStringSql(TOBRP.GetString('RESSOURCE')));
        QQ.open;
        if not QQ.Eof then
        begin
          IdRess := QQ.findField('IdRessource').asstring;
          //
          TT := TADOCommand.Create(nil);
          TT.ConnectionString := ConnectionSt;
          TRY
            Sql := 'INSERT INTO FichressProjets (IdRessource,idProjet) VALUES ('+
          					AddStringSql(idRess)+','+
          					AddStringSql(IntToStr(idProjet)) +')';
    				TT.CommandText :=Sql;
            TT.Execute;
          finally
            FreeAndNil(TT);
          end;
          //
        end;
  			QQ.Active := false;
        QQ.Close;
        FreeAndNil(QQ);
      end;
    FINALLY
    End;
  end;


var OneCmd : TADOCommand;
		ConnectionSt : string;
    Sql : string;
    IdResp : string;
    Nombase,Responsable,Libelle,CodeProjet : string;
    IdProjet : Integer;
    DateDebut,DateFin : TDateTime;
    DureeProjetJ,DureeProjetH,Cout : Double;
    Cledoc : R_CLEDOC;
    TOBressProjet : TOB;
begin
  TOBressProjet := TOB.Create ('LES RESS PROJ',nil,-1);
  //
	DecodeRefPiece(TOBTaches.GetString('CLEDOC'),cledoc);
  //
  IdProjet := TOBTaches.GetInteger('IDPROJET');
  Responsable := NormaliseStr(Copy(TOBTaches.GetString('RESPONSABLE'),1,16));
  Libelle := TOBTaches.GetString('REFERENCE');
  CodeProjet := TOBTaches.GetString('NUMDOSS');
  DateDebut := TOBTaches.GetDateTime('DATEDEBUTPROJET');
  DateFin := TOBTaches.GetDateTime('DATEFINPROJET');
  DureeProjetJ := dayspan (TOBTaches.GetDateTime('DATEDEBUTPROJET'),TOBTaches.GetDateTime('DATEFINPROJET')); // duree en jours
  DureeProjetH := HourSpan (TOBTaches.GetDateTime('DATEDEBUTPROJET'),TOBTaches.GetDateTime('DATEFINPROJET')); // duree en jours
  Cout := TOBTaches.GetDouble('COUT');
  //
  Nombase := GetOneresultSqlVP (COMMUNE,'SELECT nombaseprojet FROM FichlisteProjets WHERE IdProjet='+InttoStr(IdProjet));
  if Nombase = '' then Exit;
  TOBTaches.SetString('DATABASE',nombase);
//  ConnectionSt := ConstitueConnectionString (fServer,nombase);
  //
  ConstitueressourceProjet (Nombase,TOBTaches,TOBressProjet);

  // Association des ressources au projet
  AffecteRessourceprojet(IdProjet,TOBressProjet);
  //
	IdResp := GetOneresultSqlVP (COMMUNE,'SELECT Idressource FROM FichRessources WHERE trigramme='+AddStringSql(responsable));
  if IdResp = '' then IdResp := '-1';

 	OneCmd := TADOCommand.Create(TheForm);
  OneCmd.ConnectionString   := ConstitueConnectionString (fServer,COMMUNE);
  SQL := 'UPDATE FichListeProjets SET '+
         'NomProjet='+AddStringSql(Libelle)+','+
         'numDoss='+AddStringSql(CodeProjet)+','+
         'ID_CDP='+idResp+' '+
         'WHERE IdProjet='+InttOStr(Idprojet);
  OneCmd.CommandText := Sql;
  TRY
  	OneCmd.Execute;
  FINALLY
  	FreeAndNil(OneCmd);
  end;
  //
 	OneCmd := TADOCommand.Create(TheForm);
  ConnectionSt := ConstitueConnectionString (fServer,nombase); // base de donnéee du projet
  OneCmd.ConnectionString  := ConnectionSt;
  TRY
    if WithSupress then
    begin
      DeleteInfobase(OneCmd);
    end;
    SQL := 'UPDATE FichInfosProjet SET '+
           'NomProjet='+AddStringSql(Libelle)+','+
           'numDoss='+AddStringSql(CodeProjet)+','+
           'IdProjet='+InttoStr(IdProjet)+','+
           'IdSite=-1,'+
           'Idequipe=-1,'+
           'ID_CDP='+idResp+','+
           'Id_RTP=-1,'+
           'chaineConfigColonnesTabTaches='+AddStringSql(constitueChaineConfigColonnesTaches)+','+
           'chaineConfigGantt='+AddStringSql('0,0,1,0,0,0,0')+','+
           'chaineConfigImpression='+AddStringSql('1,0,33')+','+
           'DateHeureDebProjet='+AddDateTimeSql(DateDebut)+','+
           'DateHeureFinProjet='+AddDateTimeSql(DateFin)+','+
           'DureeProjet='+AddDoubleSql(DureeProjetJ)+','+
           'TravailProjet='+AddDoubleSql(DureeProjetH)+','+
           'TotalTravailProjet='+AddDoubleSql(DureeProjetH)+','+
           'resteAfaireProjet='+AddDoubleSql(DureeProjetH)+','+
           'CoutPrevuProjet='+AddDoubleSql(Cout)+','+
           'CoutresteAFaireProjet='+AddDoubleSql(Cout)+','+
           'DateHeureFinProjetProbable='+AddDateTimeSql(DateFin)+','+
           'realiseProjetProbable='+AddDateTimeSql(DateFin)+','+
           'CoutRealiseProjetProbable='+AddDoubleSql(Cout);
    OneCmd.CommandText := Sql;
    OneCmd.Execute;
  FINALLY
  	FreeAndNil(OneCmd);
  END;
    //
  MajDetailProjet (TOBTaches);
  ExecuteSQL('UPDATE PIECE SET GP_IDENTIFIANTWOT='+InttoStr(idProjet)+' WHERE '+WherePiece(Cledoc,ttdPiece,False));
  TOBressProjet.Free;
    //
end;

function TEchangeVP.GetServeur: string;
var iniFile : TIniFile;
		IniVp : string;
    ServeurEtPort : string;
begin
	Result := '';
  fEmplacementVp := ExtractFileDir(GetEmplacementExe);
  IniVp := IncludeTrailingBackslash(fEmplacementVp)+'visualProjet.ini';
  if not FileExists(iniVp) then Exit;
  inifile := TIniFile.Create(IniVp);
  TRY
    ServeurEtPort := iniFile.ReadString ('CONNEXION SERVEUR','nomEtPortServeur1','');
    Result := Trim(READTOKENPipe(ServeurEtPort,','));
  finally
    iniFile.Free;
  end;
end;

procedure TEchangeVP.MajDetailProjet(TOBTaches: TOB);
var ConnectionSt,Nombase : string;
		indice : Integer;
    TOBTT : TOB;
begin
  Nombase := TOBTaches.GetString('DATABASE');
  if Nombase = '' then Exit;
  ConnectionSt := ConstitueConnectionString (fServer,Nombase);
//  ConstitueRecette (ConnectionSt,TOBTaches);
  for Indice := 0 to TOBTaches.detail.count - 1 do
  begin
    TOBTT := TOBTaches.detail[Indice];
    if IsFinParagraphe(TOBTT) then Continue;
    if not isVariante (TOBTT) Then
    begin
      // les taches du projet
      EcritTache (ConnectionSt,TOBTT);
      if IsDebutParagraphe(TOBTT) then
      begin
      	ConstitueLientache(ConnectionSt,TOBTT,TOBTaches);
      end;
    end else
    begin
      // les recettes et/ou Dépenses
      if IsDebutParagraphe(TOBTT) then Continue;
//      ConstitueDepense (ConnectionSt,TOBTaches,TOBTT);;
    end;
  end;
end;

procedure TEchangeVP.EcritTache(XX: String; TOBTT: TOB);
var SQl : String;
		QQ : TADOQuery;
    TT : TADOCommand;
begin
    if TOBTT.GetInteger('GLC_IDENTIFIANTWNT') = 0 then
    begin
      // Mode Création
      SQl := ConstitueSqlInsertTache (TOBTT);
      TT := TADOCommand.create (nil);
      TT.ConnectionString := XX;
      TRY
    		TT.CommandText := Sql;
      	TT.Execute;
      FINALLY
        FreeAndNil(TT);
      END;
      //
      QQ := TADOQuery.Create(TheForm);
  		QQ.LockType := ltReadOnly;
      QQ.ConnectionString := XX;
      QQ.SQL.Add('SELECT IdTache From FichTaches WHERE numTache='+IntToStr(TOBTT.GetInteger('NUMTACHE')));
      TRY
        QQ.Open;
        if not QQ.Eof then
        begin
          TOBTT.SetInteger('GLC_IDENTIFIANTWNT',QQ.findField('IdTache').AsInteger);
          ConstitueAffectationtache (XX,QQ.findField('IdTache').AsInteger,TOBTT);
        end;
        QQ.Active := false;
        QQ.Close;
      FINALLY
        FreeAndNil(QQ);
      end;
    end else
    begin
      // Mode MAJ
    end;
end;

function TEchangeVP.ConstitueSqlInsertTache(TOBTT: TOB): string;
begin
//	TOBTT.SetDateTime('DATEFIN', AjouteDuree(TOBTT.GetDateTime('GL_DATELIVRAISON'),HeureBase100ToMinutes(TOBTT.GetDouble('MAXHRS')))); // remplacement ou nouveau
  Result := 'INSERT INTO FichTaches ('+
  'numTache,nomTache,niveauTache,estRecap,estRecapRepliee,estSupprimee,'+
  'estNonModifiable,travailAttenteSaisi,travailAttente,complements,totalTravail,'+
  'totalAttente,dateHeureDeb,dateHeureFin,dateHeureDebMiniSaisi,dateHeureDebMini,'+
  'retardAvanceSaisi,retardAvance,'+
  'margeLibre,margeLibreHeritee,realise,tauxAchevementSaisi,tauxAchevement,'+
  'resteAfaire,'+
  'codeCategorieSaisi,codeCategorie,estVisibleSurImpression,'+
  'coutTache,dureeSaisie,duree,doitRecalerTachesSuivantes,IDRecetteDepense) VALUES (';
  if IsDebutParagraphe(TOBTT) then
  begin
    Result := result +
    AddStringSql(TOBTT.GetString('NUMTACHE'))+','+   							// numtache
    AddStringSql(TOBTT.GetString('GL_LIBELLE'))+','+ 							// nomtache
    AddStringSql(TOBTT.GetString('NIVTACHE'))+',' +  							// niveautache
    AddStringSql('1')+','+ 													 							// estrecap
    AddStringSql('0')+','+ 													 							// estrecapreplie
    AddStringSql('0')+','+ 													 							// estsupprime
    AddStringSql('0')+','+ 																				// estnonmodifiable
    AddDoubleSql(0.0,' h')+','+           												// travailAttenteSaisi
    AddDoubleSql(TOBTT.GetDouble('TOTALHRS'))+','+                // travailAttente
    AddStringSql('0')+','+                                        // complements
    AddDoubleSql(TOBTT.GetDouble('MAXHRS'))+','+                  // totalTravail
    AddStringSql('0')+','+                                        // totalAttente
    AddDateTimeSql(TOBTT.GetDateTime('DATEDEBUT'))+','+           // dateHeureDeb
    AddDateTimeSql(TOBTT.GetDateTime('DATEFIN'))+','; 						// dateHeureFin
    if TOBTT.GetString('LINKOK')='-' then
    begin
      Result := result +
      AddDateTimeSql(TOBTT.GetDateTime('DATEDEBUT'))+','+           // dateHeureDeb
      AddDateTimeSql(TOBTT.GetDateTime('DATEDEBUT'))+','; 						// dateHeureFin
    end else
    begin
      Result := Result +
      AddStringSql('  /  /     00:00:00')+','+           // dateHeureDeb
      AddStringSql('  /  /     00:00:00')+','; 						// dateHeureFin
    end;
    Result := Result +
    AddStringSql('0')+','+                                        // retardAvanceSaisi
    AddStringSql('0')+','+                                        // retardAvance
    AddStringSql('0')+','+                                        // margeLibre
    AddStringSql('0')+','+                                        // margeLibreHeritee
    AddStringSql('0')+','+                                        // realise
    AddStringSql('0')+','+                                        // tauxAchevementSaisi
    AddStringSql('0')+','+                                        // tauxAchevement
    AddDoubleSql(TOBTT.GetDouble('TOTALHRS'))+','+								// resteAfaire
    AddStringSql('0')+','+                                        // codeCategorieSaisi
    AddStringSql('0')+','+                                        // codeCategorie
    AddStringSql('1')+','+                                        // estVisibleSurImpression
    AddDoubleSql(TOBTT.GetDouble('COUT'))+','+ 						// coutTache
    AddStringSql(' ')+','+                                        // dureeSaisie
    AddDoubleSql(TOBTT.GetDouble('DUREE'))+','+                   // duree
    AddStringSql('0')+','+                                        // doitRecalerTachesSuivantes
    AddStringSql('0');                                            // IDRecetteDepense
  end else
  begin
    Result := result +
    AddStringSql(TOBTT.GetString('NUMTACHE'))+','+   							// numtache
    AddStringSql(TOBTT.GetString('GL_LIBELLE'))+','+ 							// nomtache
    AddStringSql(TOBTT.GetString('NIVTACHE'))+',' +  							// niveautache
    AddStringSql('0')+','+ 													 							// estrecap
    AddStringSql('0')+','+ 													 							// estrecapreplie
    AddStringSql('0')+','+ 													 							// estsupprime
    AddStringSql('0')+','+ 																				// estnonmodifiable
    AddDoubleSql(TOBTT.GetDouble('TOTALHRS'),' h')+','+           // travailAttenteSaisi
    AddDoubleSql(TOBTT.GetDouble('TOTALHRS'))+','+                // travailAttente
    AddStringSql('0')+','+                                        // complements
    AddStringSql('0')+','+                 												// totalTravail
    AddDoubleSql(TOBTT.GetDouble('TOTALHRS'))+','+                // totalAttente
    AddDateTimeSql(TOBTT.GetDateTime('DATEDEBUT'))+','+  				  // dateHeureDeb
    AddDateTimeSql(TOBTT.GetDateTime('DATEFIN'))+','+ 						// dateHeureFin
    AddStringSql('  /  /     00:00:00')+','+           // dateHeureDebmin
    AddStringSql('  /  /     00:00:00')+','+ 						// dateHeureFin
    AddStringSql('0')+','+                                        // retardAvanceSaisi
    AddStringSql('0')+','+                                        // retardAvance
    AddStringSql('0')+','+                                        // margeLibre
    AddStringSql('0')+','+                                        // margeLibreHeritee
    AddStringSql('0')+','+                                        // realise
    AddStringSql('0')+','+                                        // tauxAchevementSaisi
    AddStringSql('0')+','+                                        // tauxAchevement
    AddDoubleSql(TOBTT.GetDouble('TOTALHRS'))+','+								// resteAfaire
    AddStringSql('0')+','+                                        // codeCategorieSaisi
    AddStringSql('0')+','+                                        // codeCategorie
    AddStringSql('1')+','+                                        // estVisibleSurImpression
    AddDoubleSql(TOBTT.GetDouble('COUT'))+','+ 						// coutTache
    AddStringSql(' ')+','+                                        // dureeSaisie
    AddDoubleSql(TOBTT.GetDouble('DUREE'))+','+                   // duree
    AddStringSql('0')+','+                                        // doitRecalerTachesSuivantes
    AddStringSql('0');                                        // IDRecetteDepense
  end;
  result := result + ')';
end;


function TEchangeVP.ConstitueAffectationtache(XX: String;Idtache: integer; TOBTT: TOB): string;
var QQ: TADOQuery;
		TT : TADOCommand;
		ConnectionSt : string;
    Sql : string;
    IdRessource : Integer;
    CodePrestation : string;
begin
	if TOBTT.GetString('GL_CODEARTICLE') = '' then Exit;
	QQ := TADOQuery.Create(TheForm);
  TRY
    ConnectionSt := ConstitueConnectionString(fServer,COMMUNE);
    QQ.ConnectionString  := ConnectionSt;
  	QQ.LockType := ltReadOnly;
    CodePrestation := NormaliseStr(Copy('_'+TOBTT.GetString('GA_LIBELLE'),1,16));
    Sql := 'SELECT IdRessource FROM FichRessources WHERE trigramme='+AddStringSql(CodePrestation);
    QQ.SQL.Add(Sql);
    QQ.Open;
    if not QQ.Eof then
    begin
      IdRessource := QQ.findField('IdRessource').AsInteger;
			SQL := 'INSERT INTO FichTachesRess (IDtache,IDressource,tauxAffectation) VALUES ('+
      AddStringSql(IntToStr(Idtache))+','+
      AddStringSql(IntToStr(IdRessource))+','+
      AddDoubleSql(100)+')';
      TT := TADOCommand.create(nil);
      TRY
        TT.ConnectionString := XX;
        TRY
    			TT.CommandText := SQL;
          TT.Execute;
        EXCEPT
          PGIInfo('Erreur ecriture affectation tache');
        end;
      FINALLY
        FreeAndNil(TT);
      END;
    end;
  FINALLY
    QQ.Active := false;
    QQ.Close;
    FreeAndNil(QQ);
  END;
end;



procedure TEchangeVP.ConstitueLientache(XX: String; TOBTT, TOBTaches: TOB);

	function 	ConstitueLientacheSql (TOBTT,TOBTaches : TOB) : string;
  var TOBP : TOB;
  begin
    result := '';
    TOBP := TOBTaches.detail[TOBTT.GetInteger('LIENTACHE')];
    if TOBP <> nil then
    begin
      result := 'INSERT INTO FichTachesPrecedentes (IDtache,IDtachePrec,IDprojet,typeLiaison) VALUES ('+
      					InttoStr(TOBTT.GetInteger('GLC_IDENTIFIANTWNT'))+','+
                InttoStr(TOBP.GetInteger('GLC_IDENTIFIANTWNT'))+','+
                InttoStr(TOBtaches.GetInteger('IDPROJET'))+','+
      					AddStringSql('FD')+
                ')';
    end;
  end;
var SQl : string;
		TT : TADOCommand;
begin
	if TOBTT.GetInteger('LIENTACHE')>= 0 then
  begin
		Sql := ConstitueLientacheSql (TOBTT,TOBTaches);
    if Sql <> '' then
    begin
      TT := TADOCommand.create (nil);
      TT.ConnectionString := XX;
      TRY
    		TT.CommandText := SQl;
      	TT.Execute;
      FINALLY
      	FreeAndNil(TT);
      end;
    end;
  end;
end;

procedure TEchangeVP.ConstitueRecette(XX : string ;TOBTaches: TOB);
var SQl : string;
		TT : TADOCommand;
begin
  if TOBTaches.Detail.Count > 0 then
  begin
    SQL := 'INSERT INTO FichRecettesDepenses (nomRecetteDepense,dateSaisie,montantRecette,'+
           'montantDepense,datePrevu,DateEffet) VALUES ('+
           AddStringSql('Recette prévue')+','+
           AddDateSql(TOBTaches.GetDateTime('DATEDEBUTPROJET'))+','+
           AddDoubleSql(TOBTaches.detail[0].GetDouble('GP_TOTALHTDEV'))+','+
           AddDoubleSql(0)+','+
           AddDateSql(TOBTaches.GetDateTime('DATEFINPROJET'))+','+
           AddDateSql(TOBTaches.GetDateTime('DATEFINPROJET'))+
           ')';
  	TT := TADOCommand.create (nil);
    TT.ConnectionString := XX;
    TRY
    	TT.CommandText := SQL;
    	TT.Execute;
    FINALLY
      FreeAndNil(TT);
    END;
  end;
end;

procedure TEchangeVP.ConstitueDepense(XX : String;TOBTaches, TOBTT: TOB);
var SQl : string;
		TT : TADOCommand;
begin
  SQL := 'INSERT INTO FichRecettesDepenses (nomRecetteDepense,dateSaisie,montantRecette,'+
  			 'montantDepense,datePrevu,DateEffet) VALUES ('+
         AddStringSql(TOBTT.GetString('GL_LIBELLE'))+','+
         AddDateSql(TOBTT.GetDateTime('DATEDEBUTPROJET'))+','+
         AddDoubleSql(0)+','+
         AddDoubleSql(TOBTT.GetDouble('GL_MONTANTPR'))+','+
         AddDateSql(TOBTaches.GetDateTime('DATEFINPROJET'))+','+
         AddDateSql(TOBTaches.GetDateTime('DATEFINPROJET'))+
         ')';
  TT := TADOCommand.Create(nil);
  TT.ConnectionString := XX;
  TRY
  	TT.CommandText := SQL;
  	TT.Execute;
  FINALLY
  	FreeAndNil(TT);
  END;
end;


procedure TEchangeVP.DeleteInfobase(CC : TADOCommand);
begin
  CC.CommandText := 'DELETE FROM FichActionsProjet';
  CC.Execute;
  //
  CC.CommandText := 'DELETE FROM FichHistoriques';
  CC.Execute;
  //
  CC.CommandText := 'DELETE FROM FichRecettesDepenses';
  CC.Execute;
  //
  CC.CommandText := 'DELETE FROM FichTaches';
  CC.Execute;
  //
  CC.CommandText := 'DELETE FROM FichTachesComplements';
  CC.Execute;
  //
  CC.CommandText := 'DELETE FROM FichTachesDetails';
  CC.Execute;
  //
  CC.CommandText := 'DELETE FROM FichTachesInterruptions';
  CC.Execute;
  //
  CC.CommandText := 'DELETE FROM FichTachesPrecedentes';
  CC.Execute;
  //
  CC.CommandText := 'DELETE FROM FichTachesReal';
  CC.Execute;
  //
  CC.CommandText := 'DELETE FROM FichTachesRess';
  CC.Execute;
  //
end;

function TEchangeVP.BaseExiste(IdBase : integer): boolean;
var QQ : TADOQuery;
		TheConnectStr : string;
begin
	Result := false;
	QQ := TADOQuery.Create(TheForm);
  TheConnectStr := ConstitueConnectionString(fServer,COMMUNE);
  QQ.ConnectionString := TheConnectStr;
  QQ.LockType := ltReadOnly;
  QQ.SQL.Add('SELECT Idprojet from FichListeProjets WHERE idProjet='+InttoStr(IdBase));
  TRY
    QQ.Open;
    Result := not QQ.Eof;
  finally
    QQ.active := False;
    QQ.Close;
    QQ.Free;
  end;
end;

procedure TEchangeVP.SetForm(const Value: Tform);
begin
	TheForm := Value;
end;

function TEchangeVP.IsliaisonPossible: Boolean;
var QQ : TADOQuery;
		ConnectionSt : string;
begin
	Result := false;
	if not Existe then
  begin
  	Exit;
  end;
  ConnectionSt := ConstitueConnectionString(fServer,COMMUNE);
  QQ := TADOQuery.create (TheForm);
  TRY
    QQ.ConnectionString := ConnectionSt;
  	QQ.LockType := ltReadOnly;
    QQ.SQL.Add('SELECT 1 from FichMetiers');
    QQ.Open;
    if not QQ.eof then
    begin
      result := true;
    end;
  FINALLY
  	QQ.Close;
    FreeAndNil(QQ);
  end;
end;

end.
