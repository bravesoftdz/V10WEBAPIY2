{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : unit de reaffectation automatique des ventilations
Suite ........ : analytiques salariés en fonction des paramètres société
Suite ........ : On doit l'utiliser en aucun de changement de stratégie de la
Suite ........ : mise en place des ventilations analytiques.
Mots clefs ... : PAIE;PGANALYTIQUE
*****************************************************************}
{ PT1 02/10/01 V562 PH on analyse le champ ANAPERSO de la table salarié pour
      ne ^pas affecté les salariés personnalisés
}
unit UTofPG_ReaffAnalAuto;

interface
uses  Windows,StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls,HTB97,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Mul,Fe_Main,DBGrids,Fiche,Qre,
{$ELSE}
       MaineAgl,eMul,eFiche,
{$ENDIF}
      Grids,HCtrls,HQry,HEnt1,EntPaie,HMsgBox,UTOF,UTOB,UTOM,Vierge,pgoutils,
      AGLInit,HStatus ;
Type
     TOF_PG_ReaffAnalAuto = Class (TOF)
       private
       LblDD, LblDF : THLabel;
       BtnLance : TToolbarButton97;
       QMul : TQUERY; // Query recuperee du mul
       procedure LanceReaffAuto (Sender: TObject);
       procedure ImprimeClick(Sender: TObject);
       procedure LanceAffectDef (Sender: TObject);
       public
       procedure OnArgument(Arguments : String ) ; override ;

     END ;

implementation

// Fonction de calcul des ventana et histoanalpaie reaffectation automatique
procedure TOF_PG_ReaffAnalAuto.LanceReaffAuto(Sender: TObject);
var St,Salarie,Etab    : String;
    TOB_Histob,TOB_VenRem,TOB_VenCot,TOBAna : TOB;
    Tob_VentilRem : TOB; // Tob des preventilations analytiques des remunerations pour recuperer le numero de compte
    Tob_VentilCot : TOB; // Tob des preventilations analytiques des Cotisations pour recuperer le numero de compte
    Pan : TPageControl;
    Tbsht : TTabSheet;
    Trace,TraceErr : TListBox;
    DateDeb,DateFin : TDateTime;
    SystemTime0 : TSystemTime;
    i : Integer;
    Q : TQuery;
begin
Salarie:=''; Etab:='';
Pan:= TPageControl (GetControl ('PANELPREP'));
Tbsht:= TTabSheet (GetControl ('TBSHTTRACE'));
Trace:= TListBox (GetControl ('LSTBXTRACE'));
TraceErr:= TListBox (GetControl ('LSTBXERROR'));
if (Trace = NIL) OR (TraceErr = NIL) then
 begin
 PGIBox ('La réaffectation ne peut pas être lancée', 'Car les composants trace ne sont pas disponibles');
 exit;
 end;
BtnLance.Enabled := FALSE;
if (Pan <> NIL) AND (Tbsht <> NIL) then Pan.ActivePage := Tbsht;
QMul.First ;
InitMove (4, ' ');
st:='SELECT * FROM VENTIL WHERE V_NATURE LIKE "RR%" ORDER BY V_NATURE,V_COMPTE';
Q:=OpenSql(st, TRUE);
TOB_VenRem := TOB.Create ('Les ventils des Rem', NIL, -1);
TOB_VenRem.LoadDetailDB('VENTIL','','',Q,FALSE,False) ;
Ferme (Q);
MoveCur (FALSE);
st:='SELECT * FROM VENTIL WHERE V_NATURE LIKE "RC%" ORDER BY V_NATURE,V_COMPTE';
Q:=OpenSql(st, TRUE);
TOB_VenCot := TOB.Create ('Les ventils des Cot', NIL, -1);
TOB_VenCot.LoadDetailDB('VENTIL','','',Q,FALSE,False) ;
Ferme (Q);
MoveCur (FALSE);
st:='SELECT * FROM VENTIREMPAIE';
Q:=OpenSql(st, TRUE);
TOB_VentilRem := TOB.Create ('Les préventils des Remunérations', NIL, -1);
TOB_VentilRem.LoadDetailDB('VENTIREMPAIE','','',Q,FALSE,False) ;
Ferme (Q);
MoveCur (FALSE);
st:='SELECT * FROM VENTICOTPAIE';
Q:=OpenSql(st, TRUE);
TOB_VentilCot := TOB.Create ('Les préventils des Cotisations', NIL, -1);
TOB_VentilCot.LoadDetailDB('VENTICOTPAIE','','',Q,FALSE,False) ;
Ferme (Q);
MoveCur (FALSE);
FiniMove ;
GetLocalTime(SystemTime0);
Trace.Items.Add('Début du traitement à :'+TimeToStr (SystemTimeToDateTime(SystemTime0)));
While Not QMul.EOF do
 begin
 Salarie:= QMul.FindField ('PPU_SALARIE').AsString;
 Etab:= QMul.FindField ('PPU_ETABLISSEMENT').AsString;
 DateDeb:=QMul.FindField ('PPU_DATEDEBUT').AsDateTime;
 DateFin:=QMul.FindField ('PPU_DATEFIN').AsDateTime;
 if TOBAna <> NIL then begin TOBAna.Free ; TOBAna:=NIL; end;
// chargement des la liste des rubriques en foction des profils
 TOB_Histob:=TOB.Create('Les lignes du bulletin',NIL,-1) ;
 st:='SELECT * FROM HISTOBULLETIN WHERE PHB_ETABLISSEMENT="'+Etab+'"'+' AND PHB_SALARIE="'+
      Salarie+'" AND PHB_DATEDEBUT="'+USDateTime(DateDeb)+'" AND PHB_DATEFIN="'+USDateTime(DateFin)+'"';
 Q:=OpenSql(st, TRUE);
 TOB_Histob.LoadDetailDB('HISTOBULLETIN','','',Q,False) ;
 ferme (Q);
 if (TOB_Histob = NIL) OR (TOB_Histob.detail.Count <= 0) then
  TraceErr.Items.Add('Pas de bulletin pour le salarié '+Salarie+' '+QMul.FindField ('PPU_LIBELLE').AsString+' '+
   QMul.FindField ('PPU_PRENOM').AsString+ ' paie du '+DateToStr (DateDeb)+ ' au '+ DateToStr (DateFin));

// Traitement de l'analytique
   TOBAna := PreVentileLignePaie ( TOB_VenRem,TOB_VenCot,TOB_Histob, Salarie, 'PRE', DateDeb, DateFin);
// Reaffectation des ventilations analytiques des cotisations en focntion de la valorisation des remunerations
   PGReaffSectionAnal (TOB_Histob,TOBAna);
 i := i + 1;
 st := 'Salarié '+Salarie+' '+QMul.FindField ('PPU_LIBELLE').AsString+' '+QMul.FindField ('PPU_PRENOM').AsString+' Etablissement : '+ RechDom ('TTETABLISSEMENT',Etab,FALSE)+' ';
 Trace.Items.Add(st);
 if (TOBAna = NIL) OR (TOBAna.detail.Count <= 0) then
  TraceErr.Items.Add('Pas de ventilation analytique pour le salarié '+Salarie+' '+QMul.FindField ('PPU_LIBELLE').AsString+' '+QMul.FindField ('PPU_PRENOM').AsString+ ' ???')
  else SauvegardeAnal(TOBAna,TOB_Histob,TOB_VentilRem,TOB_VentilCot, FALSE, Salarie, DateDeb, DateFin);

 QMul.NEXT;
 end; // Fin boucle sur la Query du mul

if TOB_Histob <> NIL then begin TOB_Histob.Free ; TOB_Histob:=NIL; end;
if TOBAna <> NIL then begin TOBAna.Free ; TOBAna:=NIL; end;
if TOB_VenRem <> NIL then begin TOB_VenRem.Free ; TOB_VenRem:=NIL; end;
if TOB_VenCot <> NIL then begin TOB_VenCot.Free ; TOB_VenCot:=NIL; end;
if TOB_VentilRem <> NIL then begin TOB_VentilRem.Free ; TOB_VentilRem:=NIL; end;
if TOB_VentilCot <> NIL then begin TOB_VentilCot.Free ; TOB_VentilCot:=NIL; end;

GetLocalTime(SystemTime0);
Trace.Items.Add('Fin du traitement à :'+TimeToStr (SystemTimeToDateTime(SystemTime0)));
Trace.Items.Add('Nombre de paies traitées : '+IntToStr (i) );
if TraceErr.Items.Count > 1 then
 begin
 Tbsht:= TTabSheet (GetControl ('TBSHTERROR'));
 if Tbsht <> NIL then Pan.ActivePage := Tbsht;
 end;
end;
// Fonction d'affectation par defaut des ventilations analytiques en fonction des paramsoc
procedure TOF_PG_ReaffAnalAuto.LanceAffectDef(Sender: TObject);
var St,Salarie       : String;
    CreationSection,Okok  : Boolean;
    AnaPerso              : String;
    Pan              : TPageControl;
    Tbsht            : TTabSheet;
    Trace,TraceErr   : TListBox;
    SystemTime0      : TSystemTime;
    i                : Integer;
    Q                : TQuery;
    TOB_Axes,TOB_Section,TSal,TS : TOB;
begin
Salarie:='';
Pan:= TPageControl (GetControl ('PANELPREP'));
Tbsht:= TTabSheet (GetControl ('TBSHTTRACE'));
Trace:= TListBox (GetControl ('LSTBXTRACE'));
TraceErr:= TListBox (GetControl ('LSTBXERROR'));
if (Trace = NIL) OR (TraceErr = NIL) then
 begin
 PGIBox ('La réaffectation ne peut pas être lancée', 'Car les composants trace ne sont pas disponibles');
 exit;
 end;
BtnLance.Enabled := FALSE;
if (Pan <> NIL) AND (Tbsht <> NIL) then Pan.ActivePage := Tbsht;
QMul.First ;
InitMove (2, ' ');
TOB_Axes := TOB.Create ('Les axes', NIL, -1);
st:='SELECT * FROM AXE';
Q:=OpenSql(st, TRUE);
TOB_Axes.LoadDetailDB('AXE','','',Q,FALSE,False) ;
Ferme (Q);
MoveCur (FALSE);
st:='SELECT S_SECTION,S_AXE FROM SECTION ORDER BY S_SECTION,S_AXE';
Q:=OpenSql(st, TRUE);
TOB_Section := TOB.Create ('Les sections', NIL, -1);
TOB_Section.LoadDetailDB('SECTION','','',Q,FALSE,False) ;
Ferme (Q);
MoveCur (FALSE);
FiniMove ;
GetLocalTime(SystemTime0);
Trace.Items.Add('Début du traitement à :'+TimeToStr (SystemTimeToDateTime(SystemTime0)));
i := 0;
CreationSection := VH_Paie.PGCreationSection;
While Not QMul.EOF do
 begin
 Salarie:= QMul.FindField ('PSA_SALARIE').AsString;
 st := 'SELECT PSA_SALARIE,PSA_ETABLISSEMENT,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_LIBREPCMB1,PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4,PSA_ANAPERSO FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"';
 Q:=OpenSql(st, TRUE);
 TSal := TOB.create ('Le Salarie', NIL, -1);
 TSal.LoadDetailDB('SALARIES','','',Q,FALSE,False) ; // chargement des infos salaries necessaires à la décomposition des sections
 Ferme (Q);
//PT1 02/10/01 V562 PH on analyse le champ ANAPERSO de la table salarié pour
 AnaPerso := 'X';
 if (TSal <> NIL) AND (TSal.detail.count >= 1) then // traitement du salarié  car la tob ne contient qu'un item
     begin
     TS := TSal.Detail [0];
     AnaPerso := TS.getvalue ('PSA_ANAPERSO');
     if AnaPerso <> 'X' then Okok := AffectVentilSalarie (TS,TOB_Axes, TOB_Section , CreationSection);
     end;
 if AnaPerso <> 'X' then
  begin
  if (Okok) then
    begin
    i := i +1;
    Trace.Items.Add('Ventilations analytiques générées pour le salarié '+Salarie+' '+QMul.FindField ('PSA_LIBELLE').AsString+' '+QMul.FindField ('PSA_PRENOM').AsString)
    end
    else   TraceErr.Items.Add('Pas de ventilation analytique générée pour le salarié '+Salarie+' '+QMul.FindField ('PSA_LIBELLE').AsString+' '+QMul.FindField ('PSA_PRENOM').AsString+ ' ???');
  end;
// FIN PT1
 if TSal <> NIL then begin TSal.Free ; TSal:=NIL; end;
 QMul.Next;
 end;

if TOB_Axes <> NIL then begin TOB_Axes.Free ; TOB_Axes:=NIL; end;
if TOB_Section <> NIL then begin TOB_Section.Free ; TOB_Section:=NIL; end;
if TSal <> NIL then begin TSal.Free ; TSal:=NIL; end;

GetLocalTime(SystemTime0);
Trace.Items.Add('Fin du traitement à :'+TimeToStr (SystemTimeToDateTime(SystemTime0)));
Trace.Items.Add('Nombre de salariés traités : '+IntToStr (i) );
if TraceErr.Items.Count > 1 then
 begin
 Tbsht:= TTabSheet (GetControl ('TBSHTERROR'));
 if Tbsht <> NIL then Pan.ActivePage := Tbsht;
 end;

end;

procedure TOF_PG_ReaffAnalAuto.ImprimeClick(Sender: TObject);
var
MPages:tpagecontrol;
begin
{$IFNDEF EAGLCLIENT}
MPages := TPageControl(getcontrol('PANELPREP'));
if MPages <> nil then
PrintPageDeGarde(MPages,TRUE,TRUE,FALSE,Ecran.Caption,0) ;
{$ENDIF}
end;

procedure TOF_PG_ReaffAnalAuto.OnArgument(Arguments: String);
var F : TFVierge ;
    st : STring;
    BImprime : ttoolbarbutton97;
    TbshtDefinition,TbshtAffect  : TTabSheet;
begin
inherited ;
st:= Trim (Arguments);
if St = 'Y' then Ecran.Caption := 'Affectation automatique des sections analytiques salariés';
UpdateCaption(TFVierge(Ecran));

TbshtDefinition:= TTabSheet (GetControl ('TBSHTDEFINITION'));
TbshtAffect := TTabSheet (GetControl ('TBSHTAFFECT'));
if st = 'Y' then
   begin
   if TbshtDefinition <> NIL then TbshtDefinition.TabVisible := FALSE;
   if TbshtAffect <> NIL then TbshtAffect.TabVisible := TRUE;
   end
   else
   begin
   if TbshtDefinition <> NIL then TbshtDefinition.TabVisible := TRUE;
   if TbshtAffect <> NIL then TbshtAffect.TabVisible := FALSE;
   end;

if st <> 'Y' then
 begin
 LblDD:=THLabel (GetControl ('DATEDEBUT'));
 LblDF:=THLabel (GetControl ('DATEFIN'));
 if LblDD <> NIL then LblDD.Caption := ReadTokenSt(st);
 if LblDF <> NIL then LblDF.Caption := ReadTokenSt(st);
 end;
BtnLance:=TToolbarButton97 (GetControl ('BTNLANCE'));
if BtnLance<>NIL then
 begin
 if St <> 'Y'  then BtnLance.OnClick := LanceReaffAuto
  else BtnLance.OnClick := LanceAffectDef;
 end;
BImprime :=  ttoolbarbutton97(getcontrol('BIMPRIMER'));
if Bimprime <> nil then  Bimprime.Onclick := ImprimeClick;
if not (Ecran is TFVierge) then exit;
F:=TFVierge(Ecran) ;
{$IFDEF EAGLCLIENT}
   QMUL:=THQuery(F.FMULQ).TQ ;
{$ELSE}
   QMUL:=F.FMULQ ;
{$ENDIF}
end;

Initialization
registerclasses([TOF_PG_ReaffAnalAuto]);
end.
