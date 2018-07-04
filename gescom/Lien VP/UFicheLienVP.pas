unit UFicheLienVP;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Dialogs, HTB97, ExtCtrls, HPanel,Grids, Hctrls,UTOB,UEchangeVP,HEnt1,Paramsoc,HmsgBox,
  DateUtils, StdCtrls, TntExtCtrls;

type

  TFFicheTestVP = class(TForm)
    PBAS: THPanel;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    ImgHrs: THImage;
    ImgMetiers: THImage;
    ImgRessources: THImage;
    Label4: TLabel;
    ImgConges: THImage;
    Label5: TLabel;
    ImgInterv: THImage;
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Déclarations privées }
    TheEchangeVp : TEchangeVP;
    procedure ConstitueHoraires (TOBXX : TOB);
    procedure DefiniMetiers(TOBXX: TOB);
    procedure DefiniRessources(TOBXX: TOB);
    procedure DefiniTypeRessource(TOBXX: TOB);
//    procedure DefiniPrestations(TOBXX: TOB);
    procedure DefiniCongesRessources(TOBXX: TOB);
    procedure DefiniAffectInterv (TOBXX : TOB);

  public
    { Déclarations publiques }
  end;


procedure AppelLienVP;
procedure AppelVP;

implementation

{$R *.dfm}

procedure AppelLienVP;
var XX : TFFicheTestVP;
begin
	XX := TFFicheTestVP.create(Application);
  Try
  	XX.ShowModal ;
  finally
    XX.Free;
  end;
end;

procedure AppelVP;
var CurDir,VPDir,EXEVP : string;
begin
  Curdir := ExtractFileDir(Application.exename);
  EXEVP := GetEmplacementVPExe;
  if EXEVP = '' then exit;
  VPDir := ExtractFileDir(EXEVP);
  ChDir(VPDIR);
	FileExec (EXEVP,False,false);
  ChDir(CurDir);
end;


{ TFFicheTestVP }

procedure TFFicheTestVP.ConstitueHoraires(TOBXX: TOB);
var TOBH : TOB;
begin
	TOBH := TOB.Create('HORAIRES STD',TOBXX,-1);
  TOBH.AddChampSupValeur('DEBAM',GetparamsocSecur('SO_BTAMDEBUT',StrToTime('08:00')));
  TOBH.AddChampSupValeur('FINAM',GetparamsocSecur('SO_BTAMFIN',StrToTime('12:00')));
  TOBH.AddChampSupValeur('DEBUTPM',GetparamsocSecur('SO_BTPMDEBUT',StrToTime('14:00')));
  TOBH.AddChampSupValeur('FINPM',GetparamsocSecur('SO_BTPMFIN',StrToTime('18:00')));
  TheEchangeVp.SetHorairesSTD(TOBXX);
  ImgHrs.Enabled := True;
  ImgHrs.Visible := True;
end;


procedure TFFicheTestVP.DefiniMetiers (TOBXX : TOB);
var QQ : Tquery;
begin
  QQ := OpenSQL('SELECT * FROM NATUREPREST WHERE BNP_TYPERESSOURCE IN ("SAL","MAT")',True,-1,'',true);
  TRY
    if Not QQ.eof then
    begin
      TOBXX.LoadDetailDB('NATUREPREST','','',QQ,false);  // les metiers
      TheEchangeVp.SetMetiers(TOBXX);
    end;
  finally
    Ferme(QQ);
  end;
end;

procedure TFFicheTestVP.DefiniTypeRessource (TOBXX : TOB);
var QQ : Tquery;
begin
  QQ := OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="TRE" AND CC_CODE IN ("SAL","MAT")',True,-1,'',true); //
  TRY
    if Not QQ.eof then
    begin
      TOBXX.LoadDetailDB('CHOIXCOD','','',QQ,false);  // les metiers
      TheEchangeVp.SetTypeRessource(TOBXX);
    end;
    ImgMetiers.Enabled := True;
    ImgMetiers.Visible := True;
  finally
    Ferme(QQ);
  end;
end;

procedure TFFicheTestVP.DefiniRessources (TOBXX :TOB);
var QQ : Tquery;
    StSQL : string;
begin
(*
  StSql := 'SELECT ARS_RESSOURCE, ARS_TYPERESSOURCE,ARS_ARTICLE,ARS_LIBELLE, ARS_LIBELLE2, '+
           'ARS_TAUXREVIENTUN,'+
           '(SELECT CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="TRE" AND CC_CODE=ARS_TYPERESSOURCE) AS BNP_LIBELLE,'+
           '(SELECT BNP_LIBELLE FROM NATUREPREST WHERE BNP_NATUREPRES='+
           '(SELECT GA_NATUREPRES FROM ARTICLE WHERE GA_ARTICLE=ARS_ARTICLE)) AS BNP_LIBELLE2 '+
            'FROM RESSOURCE WHERE ARS_TYPERESSOURCE IN ("SAL","MAT") AND '+
            'ARS_FERME="-" ORDER BY ARS_TYPERESSOURCE';
*)
  StSql := 'SELECT ARS_RESSOURCE, ARS_TYPERESSOURCE,ARS_ARTICLE,ARS_LIBELLE, ARS_LIBELLE2, '+
           'ARS_TAUXREVIENTUN,'+
           '(SELECT CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="TRE" AND CC_CODE=ARS_TYPERESSOURCE) AS BNP_LIBELLE,'+
           '(SELECT BNP_LIBELLE FROM NATUREPREST WHERE BNP_NATUREPRES=(SELECT GA_NATUREPRES FROM ARTICLE WHERE GA_ARTICLE=ARS_ARTICLE)) AS BNP_LIBELLE2, '+
           'METIER = CASE '+
           'WHEN (SELECT BNP_LIBELLE FROM NATUREPREST WHERE BNP_NATUREPRES=(SELECT GA_NATUREPRES FROM ARTICLE WHERE GA_ARTICLE=ARS_ARTICLE)) IS NULL THEN (SELECT CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="TRE" AND CC_CODE=ARS_TYPERESSOURCE) '+
           'ELSE (SELECT BNP_LIBELLE FROM NATUREPREST WHERE BNP_NATUREPRES=(SELECT GA_NATUREPRES FROM ARTICLE WHERE GA_ARTICLE=ARS_ARTICLE)) '+
           'END '+
           'FROM RESSOURCE WHERE ARS_TYPERESSOURCE IN ("SAL","MAT") AND '+
//           '(ARS_RESSOURCE="0000000079") AND ' + 
           'ARS_FERME="-" ORDER BY METIER';
  QQ := OPENSQL(STSQL, False);
  TRY
    if Not QQ.Eof then
    begin
      TOBXX.LoadDetailDB('RESSOURCE', '', '', QQ,false); // les ressources
      TheEchangeVp.SetRessources(TOBXX);
    end;
    ImgRessources.Enabled := True;
    ImgRessources.Visible := True;
  FINALLY
    ferme(QQ);
  end;
end;

procedure TFFicheTestVP.DefiniCongesRessources (TOBXX : TOB);

  function ConstitueRequeteAbsenceConges (DateD,DateF : TdateTime) : string ;
  begin
    result := '( '+ // 1
              '(PCN_DATEDEBUTABS>="' + UsDateTime(DateD) + '" AND PCN_DATEDEBUTABS<="' + UsDateTime(DateF) + '") OR '+
              '(PCN_DATEFINABS>="' + USDateTime(DateD) + '" AND PCN_DATEFINABS<="' + USDateTime(DateF) + '") OR ' +
              '(PCN_DATEDEBUTABS<="' + UsDateTime(DateD) + '" AND PCN_DATEFINABS>="' + UsDateTime(DateF) + '")' +
              ')'; //1
  end;

var StSql : string;
		DateDebut,DateFin : TDateTime;
    QQ : TQuery;
begin
  // chargement des événements absences de la paie pour les salariés dans la période
  DateDebut :=  V_PGI.DATEENTREE;
  DateFin   := IncWeek (V_PGI.DATEENTREE,5);
 	StSQL := 'SELECT *,R1.ARS_RESSOURCE,R1.ARS_LIBELLE2,R1.ARS_LIBELLE FROM ABSENCESALARIE '+
					 'LEFT JOIN RESSOURCE R1 ON R1.ARS_SALARIE=PCN_SALARIE '+
					 'WHERE (R1.ARS_RESSOURCE <>'') AND ('+ ConstitueRequeteAbsenceConges(DateDebut,DateFin) +') AND '+
           '(PCN_DATEANNULATION="'+USDATETIME(Idate1900)+'") AND '+
           '(PCN_VALIDRESP<>"") AND (R1.ARS_FERME="-")';
	QQ := OpenSql (StSQL,true,-1,'',true);
  if not QQ.eof then
  begin
  	TobXX.loadDetailDb('ABSENCESALARIE','','', QQ,false);
    TheEchangeVp.SetAbsencesSalaries(TOBXX,DateDebut,DateFin);
  end;
  ImgConges.Enabled := True;
  ImgConges.Visible := True;
  ferme(QQ);

end;

(*
procedure TFFicheTestVP.DefiniPrestations(TOBXX: TOB);
var QQ : Tquery;
    StSQL : string;
begin
  StSQL := 'SELECT GA_CODEARTICLE,GA_LIBELLE,GA_NATUREPRES,GA_DPR,'+ // Element de la prestation (ressource virtuelle)
           'N.BNP_TYPERESSOURCE,N.BNP_LIBELLE '+ // element de la ressource
           'FROM ARTICLE '+
           'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=GA_NATUREPRES '+
           'WHERE GA_TYPEARTICLE IN ("FRA","PRE") AND N.BNP_TYPERESSOURCE IN ("SAL","MAT")';
//           'WHERE GA_TYPEARTICLE IN ("FRA","PRE") AND N.BNP_TYPERESSOURCE IN ("SAL","INT","MAT","OUT")';
  QQ := OPENSQL(STSQL, False);
  TRY
    if Not QQ.Eof then
    begin
      TOBXX.LoadDetailDB('ARTICLE', '', '', QQ,false); // les ressources
      TheEchangeVp.SetRessourcesVirt(TOBXX);
    end;
    ImgRessources.Enabled := True;
    ImgRessources.Visible := True;
  FINALLY
    ferme(QQ);
  end;
end;
*)

procedure TFFicheTestVP.BValiderClick(Sender: TObject);
var TOBXX : TOB;
begin
	TOBXX := TOB.Create ('LES DATAS',nil,-1);
  TRY
    //
    ConstitueHoraires (TOBXX);
    TOBXX.ClearDetail;
    //
    DefiniMetiers (TOBXX);
    TOBXX.ClearDetail;
    //
    DefiniTypeRessource (TOBXX);
    TOBXX.ClearDetail;
    //
    DefiniRessources (TOBXX);
    TOBXX.ClearDetail;
    //
    (*
    DefiniPrestations (TOBXX);
    TOBXX.ClearDetail;
    *)
    //
    DefiniCongesRessources (TOBXX);
    TOBXX.ClearDetail;
    //
    DefiniAffectInterv (TOBXX);
    TOBXX.ClearDetail;

    PgiInfo ('Traitement Terminé....');
  FINALLY
    TOBXX.Free;
    close;
  end;
end;

procedure TFFicheTestVP.DefiniAffectInterv(TOBXX: TOB);

  function ConstitueRequeteInterv (DateD,DateF : TdateTime) : string ;
  begin
    result := '( '+ // 1
              '(BEP_DATEDEB>="' + UsDateTime(DateD) + '" AND BEP_DATEDEB<="' + UsDateTime(DateF) + '") OR '+
              '(BEP_DATEFIN>="' + USDateTime(DateD) + '" AND BEP_DATEFIN<="' + USDateTime(DateF) + '") OR ' +
              '(BEP_DATEDEB<="' + UsDateTime(DateD) + '" AND BEP_DATEFIN>="' + UsDateTime(DateF) + '")' +
              ')'; //1
  end;

var StSql : string;
		DateDebut,DateFin : TDateTime;
    QQ : TQuery;
begin
  // chargement des événements concernant les interventions sur la période
  DateDebut :=  V_PGI.DATEENTREE;
  DateFin   := IncWeek (V_PGI.DATEENTREE,6);
 	StSQL := 'SELECT * FROM BTEVENPLAN WHERE '+ ConstitueRequeteInterv(DateDebut,DateFin);
	QQ := OpenSql (StSQL,true,-1,'',true);
  if not QQ.eof then
  begin
  	TobXX.loadDetailDb('BTEVENPLAN','','', QQ,false);
    TheEchangeVp.SetInterv(TOBXX,DateDebut,DateFin);
  end;
  ImgInterv.Enabled := True;
  ImgInterv.Visible := True;
  ferme(QQ);
  // chargement des taches affectés aux ressources
  // TODO
  // --------------
end;

procedure TFFicheTestVP.FormShow(Sender: TObject);
begin
	if not TheEchangeVp.Existe then
  begin
  	BValider.Enabled := false;
    PGIInfo('VisualProjet n''est pas installé sur ce poste');
  end;
end;

procedure TFFicheTestVP.FormCreate(Sender: TObject);
begin
  TheEchangeVp := TEchangeVP.create;
  TheEchangeVp.ecran := Self;
end;

procedure TFFicheTestVP.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
	if TheEchangeVp <> nil then TheEchangeVp.free;
end;

end.
