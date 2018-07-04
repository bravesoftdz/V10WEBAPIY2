unit UtofPlanningGlo;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}

{$ELSE}
   dbTables, db,HDB,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOF, vierge,AffaireUtil,UTob,Grids,EntGC,
      DicoAF,Saisutil,Hstatus,M3FP,HPlanning;
Type
     TOF_PlanningGlo = Class (TOF)
     public
        procedure OnLoad ; override ;
        procedure OnClose ; override ;

     private
     TOBRES,TOBACT,TobEtats : TOB;
       procedure AlimPlanning;
    END ;
implementation


procedure TOF_PlanningGlo.OnLoad;
Begin
 Inherited;

End;

procedure TOF_PlanningGlo.OnClose;

Begin
 Inherited;
TOBRES.Free;
TOBACT.Free;
TobEtats.Free;
End;

procedure TOF_PlanningGlo.AlimPlanning;
var Pl : THPlanning;
    CodeRes,stSQL,ResEnc : String;
    DateDeb : TDateTime;
    TOBDetItems, TobdetRes,tobDetEtat : TOB;
    Q : TQuery;
    i : integer;
Begin
 Inherited;
// Recup param fiche AGL
Pl := THPlanning( GetControl('PLANNINGACT'));
DateDeb := StrToDate(GetControlText('ACT_DATEACTIVITE'));

// Paramètres généraux
//Pl.DateOfStart := DateDeb;

// **************  Chargement Tob items  *********************
stSQL := 'SELECT ATA_DATEDEBUT,ATA_DATEFIN,ATA_LIBELLE,ATA_AFFAIRE,ATA_NUMEROTACHE,ATA_RESSOURCE' +
      ' FROM TACHE WHERE ATA_RESSOURCE<>"" AND ATA_DATEDEBUT>="'+usdateTime(DateDeb)+ '"' ;
Q := OPENSQL(stSQL,True);
TOBACT := TOB.Create('liste des items',nil,-1);
TOBACT.LoadDetailDB('SELECT tache','','',Q,False);
ferme(Q);

// Correspondance des champs
Pl.ChampdateDebut:='ATA_DATEDEBUT';
Pl.ChampDateFin:= 'ATA_DATEFIN';
Pl.ChampLibelle:='ATA_LIBELLE';
Pl.ChampEtat:='ATA_AFFAIRE';
Pl.ChampLineID:='ATA_RESSOURCE';
Pl.ChampHint :='ATA_LIBELLE';
Pl.TOBItems := TOBACT;

// **************** Chargement TOB Ressource  **********************
TOBRes := TOB.Create ('Liste des ressources',Nil,-1);
stSQL := 'SELECT ARS_RESSOURCE,ARS_LIBELLE from RESSOURCE' ;
Q := OPENSQL(stSQL,True);
TOBRes.LoadDetailDB('SELECT ressource','','',Q,False);
ferme(Q);

Pl.TobRes := TobRes;
Pl.ResChampID := 'ARS_RESSOURCE';
PL.TokenFieldColFixed := 'ARS_LIBELLE' ;
PL.TokenSizeColFixed  := '120' ;
PL.TokenAlignColFixed := 'L' ;

// ****************** Chargement des Etats **************************
TobEtats := Tob.Create ('liste des etats', Nil,-1);
stSQL := 'SELECT AFF_AFFAIRE,AFF_AFFAIRE1,AFF_AFFAIRE2,AFF_AFFAIRE3, AFF_LIBELLE FROM AFFAIRE';
Q := OPENSQL(stSQL,True);
TOBEtats.LoadDetailDB('SELECT etats','','',Q,False);
ferme(Q);
if TOBEtats.Detail.Count > 0 then
    BEGIN
    TobDetEtat :=TOBEtats.Detail[0];
    TobDetEtat.AddchampSup ('BGC',True);
    TobDetEtat.AddchampSup ('COLOR',True);
    TobDetEtat.AddchampSup ('FONTE',True);
    for i := 0 to  TOBEtats.Detail.Count-1 do
        BEGIN
        TobDetEtat :=TOBEtats.Detail[i];
        TobDetEtat.putvalue ('BGC','clBlue');
        TobDetEtat.putvalue ('COLOR','$0080FFFF');
        TobDetEtat.putvalue ('FONTE','MS Sans Serif');
        END;
    END;

PL.EtatChampCode  := 'AFF_AFFAIRE' ;
PL.EtatChampLibelle := 'AFF_LIBELLE' ;
PL.EtatChampBackGroundColor := 'BGC' ;
PL.EtatChampFontColor := 'COLOR' ;
PL.EtatChampFontName := 'FONTE' ;
PL.TobEtats := TobEtats ;

Pl.IntervalDebut := DateDeb; //StrToDateTime ('01/11/2000') ;
Pl.IntervalFin := plusdate(DateDeb,3,'M');    //StrToDateTime ('01/04/2001') ;

Pl.Activate := True;


End;

// Fonctions AGL
procedure AGLAlimPlanningGlo(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;
if (MaTOF is TOF_PlanningGlo) then TOF_PlanningGlo(MaTOF).AlimPlanning() else exit;
end;

Initialization
registerclasses([TOF_PlanningGlo]);
RegisterAglProc('AlimPlanningGlo',TRUE,0,AGLAlimPlanningGlo);
end.
