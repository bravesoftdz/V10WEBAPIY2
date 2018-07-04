{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/09/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : DADSU_ETAB ()
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
{
PT1   : 12/01/2007 VG V_80 Ajout du total base formation CDD - FQ N°15102
}
Unit UTOFPG_DADSU_ETAB;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     DBCtrls,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$else}
     UtileAGL,
{$ENDIF}
     uTob,
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
     HTB97,
     Vierge,
     PgDADSOutils,
     PgDADSCommun,
     P5Def,
     windows;

Type
  TOF_DADSU_ETAB = Class (TOF)
  public
    procedure OnArgument (S : String ) ; override;
    procedure OnLoad                   ; override;
    procedure OnClose                  ; override;
    procedure OnNew                    ; override;
    procedure OnDelete                 ; override;
    procedure OnUpdate                 ; override;
    procedure OnDisplay                ; override;
    procedure OnCancel                 ; override;
  private

    Etab, State : string;

    QMul : TQUERY;     // Query recuperee du mul

    TAssujetti, PAssujetti : TCheckBox;

    T_Periode : TOB;

    ControleOK, EtabChange : Boolean;

    Daccord, EtabPrem, EtabPrec, EtabSuiv, EtabDern : TToolBarButton97;

    procedure AfficheCaption();
    procedure ChargeZones ();
    procedure SauveZones();
    procedure MetABlanc();
    function ControleConform() : boolean;
    function UpdateTable() : boolean;
    procedure AssujettiClick(Sender: TObject);
{$IFNDEF EAGLCLIENT}
    procedure EtabPremClick (Sender: TObject);
    procedure EtabPrecClick (Sender: TObject);
    procedure EtabSuivClick (Sender: TObject);
    procedure EtabDernClick (Sender: TObject);
{$ENDIF}
    Function BougeEtab(Button: TNavigateBtn) : boolean ;
    procedure GereQueryEtab();
    procedure Impression(Sender: TObject);
    procedure Validation(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  end;

Implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/09/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_DADSU_ETAB.OnArgument (S : String );
var
Pages : TPageControl;
Arg : string;
begin
Inherited;
Arg:= S;
State:= Trim (ReadTokenPipe (Arg, ';'));
Etab:= Trim (ReadTokenPipe (Arg, ';'));
TypeD:= Trim (ReadTokenPipe (Arg, ';'));

TFVierge(Ecran).OnKeyDown:=FormKeyDown;
TAssujetti:= TCheckBox (GetControl ('CHTASSUJETTI'));
if (TAssujetti<>nil) then
   TAssujetti.OnClick:= AssujettiClick;
PAssujetti:= TCheckBox (GetControl ('CHPASSUJETTI'));
if (PAssujetti<>nil) then
   PAssujetti.OnClick:= AssujettiClick;

SetControlText('EETAB', Etab);

// Positionnement sur le premier onglet
Pages := TPageControl(GetControl('PAGES'));
if Pages<>nil then
   Pages.ActivePageIndex:=0;

// recuperation de la query du multicritere
QMul := TFVierge(Ecran).FMULQ ;

// Gestion du navigateur
EtabPrem:= TToolbarButton97 (GetControl ('BETABPREM'));
EtabPrec:= TToolbarButton97 (GetControl ('BETABPREC'));
EtabSuiv:= TToolbarButton97 (GetControl ('BETABSUIV'));
EtabDern:= TToolbarButton97 (GetControl ('BETABDERN'));
{$IFNDEF EAGLCLIENT}
if (EtabPrem<>NIL) then
   begin
   EtabPrem.Enabled:= True;
   EtabPrem.Visible:= True;
   EtabPrem.OnClick:= EtabPremClick;
   end;

if (EtabPrec<>NIL) then
   begin
   EtabPrec.Enabled:= True;
   EtabPrec.Visible:= True;
   EtabPrec.OnClick:= EtabPrecClick;
   end;

if (EtabSuiv<>NIL) then
   begin
   EtabSuiv.Enabled:= True;
   EtabSuiv.Visible:= True;
   EtabSuiv.OnClick:= EtabSuivClick;
   end;

if (EtabDern<>NIL) then
   begin
   EtabDern.Enabled:= True;
   EtabDern.Visible:= True;
   EtabDern.OnClick:= EtabDernClick;
   end;
{$ELSE}
if (EtabPrem<>NIL) then
   EtabPrem.Visible:= False;

if (EtabPrec<>NIL) then
   EtabPrec.Visible:= False;

if (EtabSuiv<>NIL) then
   EtabSuiv.Visible:= False;

if (EtabDern<>NIL) then
   EtabDern.Visible:= False;
{$ENDIF}

TFVierge (Ecran).BImprimer.OnClick:= Impression;

Daccord:= TToolbarButton97 (GetControl ('BDACCORD'));
if (Daccord<>NIL) then
   begin
   Daccord.Enabled:= True;
   Daccord.Visible:= True;
   Daccord.OnClick:= Validation;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/09/2007
Modifié le ... :   /  /    
Description .. : Procédure exécutée sur le chargement de la fiche
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_DADSU_ETAB.OnLoad;
begin
Inherited;
ChargeZones;
AfficheCaption;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/09/2007
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_DADSU_ETAB.OnClose;
begin
Inherited;
end;

procedure TOF_DADSU_ETAB.OnNew;
begin
Inherited;
end;

procedure TOF_DADSU_ETAB.OnDelete;
begin
Inherited;
end;

procedure TOF_DADSU_ETAB.OnUpdate;
begin
Inherited;
end;

procedure TOF_DADSU_ETAB.OnDisplay ();
begin
Inherited;
end;

procedure TOF_DADSU_ETAB.OnCancel ();
begin
Inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/09/2007
Modifié le ... :   /  /    
Description .. : Procédure exécutée sur le chargement de la fiche
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_DADSU_ETAB.AfficheCaption();
var
EtabCaption : string;
begin
EtabCaption:= 'Etablissement '+GetControlText ('EETAB')+' '+
              GetControlText ('ELIBELLE');
Ecran.Caption:= EtabCaption;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/09/2007
Modifié le ... :
Description .. : Chargement des éléments de la fiche
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_DADSU_ETAB.ChargeZones();
var
StDADSD, StEtab : String;
QRechDADSD, QRechEtab : TQuery;
TDetail, TDetailD : Tob;
begin
MetABlanc;
SetControlText('EETAB', '');
SetControlText('ELIBELLE', '');



StEtab:= 'SELECT ET_LIBELLE'+
        ' FROM ETABLISS WHERE'+
        ' ET_ETABLISSEMENT="'+Etab+'"';

QRechEtab:= OpenSql (StEtab, TRUE);
if (not QRechEtab.EOF) then
   begin
   SetControlText ('EETAB', Etab);
   SetControlText ('ELIBELLE', QRechEtab.FindField ('ET_LIBELLE').Asstring);

   StDADSD:= 'SELECT PDS_SALARIE, PDS_TYPE, PDS_ORDRE, PDS_DATEDEBUT,'+
             ' PDS_DATEFIN, PDS_ORDRESEG, PDS_SEGMENT, PDS_DONNEE,'+
             ' PDS_DONNEEAFFICH'+
             ' FROM DADSDETAIL WHERE'+
             ' PDS_SALARIE="**'+Etab+'" AND'+
             ' PDS_TYPE="'+TypeD+'" AND'+
             ' PDS_ORDRE=2 AND'+
             ' PDS_EXERCICEDADS = "'+PGExercice+'"'+
             ' ORDER BY PDS_ORDRESEG,PDS_SEGMENT,PDS_DATEDEBUT,'+
             ' PDS_DATEFIN';
   QRechDADSD:= OpenSql (StDADSD,TRUE);
   TDetail:= TOB.Create ('Les détails', NIL, -1);
   TDetail.LoadDetailDB ('DADSDETAIL', '', '', QRechDADSD, False);
   Ferme (QRechDADSD);

   TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'], ['S80.G62.05.001'], TRUE);
   if ((TAssujetti<>NIL) and (TDetailD<>NIL)) then
      begin
      TAssujetti.Checked:= TDetailD.GetValue ('PDS_DONNEEAFFICH')='01';
      SetControlEnabled ('LTAXE', True);
      SetControlEnabled ('ETAXE', True);

      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'],
                                    ['S80.G62.05.002.001'], TRUE);
      if (TDetailD<>NIL) then
         SetControlText ('ETAXE', TDetailD.GetValue ('PDS_DONNEEAFFICH'));
      end
   else
      begin
      SetControlEnabled ('LTAXE', False);
      SetControlEnabled ('ETAXE', False);
      end;

   TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'], ['S80.G62.10.001'], TRUE);
   if ((PAssujetti<>NIL) and (TDetailD<>NIL)) then
      begin
      PAssujetti.Checked:= TDetailD.GetValue ('PDS_DONNEEAFFICH')='01';
//PT1
      SetControlEnabled ('LFORMATIONCDD', True);
      SetControlEnabled ('EFORMATIONCDD', True);
//FIN PT1
      SetControlEnabled ('LFORMATION', True);
      SetControlEnabled ('EFORMATION', True);

//PT1
      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'],
                                    ['S80.G62.10.002.001'], TRUE);
      if (TDetailD<>NIL) then
         SetControlText ('EFORMATIONCDD', TDetailD.GetValue ('PDS_DONNEEAFFICH'));
//FIN PT1

      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'],
                                    ['S80.G62.10.003.001'], TRUE);
      if (TDetailD<>NIL) then
         SetControlText ('EFORMATION', TDetailD.GetValue ('PDS_DONNEEAFFICH'));
      end
   else
      begin
//PT1
      SetControlEnabled ('LFORMATIONCDD', False);
      SetControlEnabled ('EFORMATIONCDD', False);
//FIN PT1
      SetControlEnabled ('LFORMATION', False);
      SetControlEnabled ('EFORMATION', False);
      end;
   end;
Ferme (QRechEtab);
if (TDetail<>nil) then
   FreeAndNil (TDetail);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 25/09/2007
Modifié le ... :
Description .. : Enregistrement des éléments de la fiche
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_DADSU_ETAB.SauveZones();
var
CleDADS : TCleDADS;
begin
Trace:= TStringList.Create;

CleDADS.Salarie:= '**'+Etab;
CleDADS.TypeD:= TypeD;
CleDADS.Num:= 2;
CleDADS.DateDeb:= DebExer;
CleDADS.DateFin:= FinExer;
CleDADS.Exercice:= PGExercice;
if (TAssujetti.Checked=True) then
   begin
   CreeDetail (CleDADS, 71, 'S80.G62.05.001', '01', '01');
//PT1
   if (GetControlText ('ETAXE')='') then
      SetControlText ('ETAXE', '0');
//FIN PT1
   CreeDetail (CleDADS, 72, 'S80.G62.05.002.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('ETAXE')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('ETAXE')), 0)));
   if (StrToFloat (GetControlText ('ETAXE'))<0) then
      CreeDetail (CleDADS, 73, 'S80.G62.05.002.002', 'N', 'N');
   end
else
   CreeDetail (CleDADS, 71, 'S80.G62.05.001', '02', '02');

if (PAssujetti.Checked=True) then
   begin
   CreeDetail (CleDADS, 81, 'S80.G62.10.001', '01', '01');
//PT1
   if (GetControlText ('EFORMATIONCDD')='') then
      SetControlText ('EFORMATIONCDD', '0');
   CreeDetail (CleDADS, 82, 'S80.G62.10.002.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EFORMATIONCDD')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EFORMATIONCDD')), 0)));
   if (StrToFloat (GetControlText ('EFORMATIONCDD'))<0) then
      CreeDetail (CleDADS, 83, 'S80.G62.10.002.002', 'N', 'N');

   if (GetControlText ('EFORMATION')='') then
      SetControlText ('EFORMATION', '0');
//FIN PT1
   CreeDetail (CleDADS, 84, 'S80.G62.10.003.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EFORMATION')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EFORMATION')), 0)));
   if (StrToFloat (GetControlText ('EFORMATION'))<0) then
      CreeDetail (CleDADS, 85, 'S80.G62.10.003.002', 'N', 'N');
   end
else
   CreeDetail (CleDADS, 81, 'S80.G62.10.001', '02', '02');

{$IFNDEF DADSUSEULE}
CreeJnalEvt ('001', '047', 'OK', NIL, NIL, Trace);
{$ENDIF}
if Trace <> nil then
   FreeAndNil (Trace);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 25/09/2007
Modifié le ... :   /  /
Description .. : Effacement des zones de la fiche de saisie établissement
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_DADSU_ETAB.MetABlanc();
begin
if (TAssujetti<>NIL) then
   TAssujetti.Checked:= False;
SetControlText('ETAXE', '');
if (PAssujetti<>nil) then
   PAssujetti.Checked:= False;
SetControlText('EFORMATIONCDD', ''); //PT1
SetControlText('EFORMATION', '');
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 25/09/2007
Modifié le ... :   /  /
Description .. : Contrôle des données
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
function TOF_DADSU_ETAB.ControleConform() : boolean;
begin
result:= TRUE;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 25/09/2007
Modifié le ... :   /  /
Description .. : Mise à jour des tables
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
function TOF_DADSU_ETAB.UpdateTable(): boolean;
var
Rep : integer;
begin
result := FALSE;

Rep:=PGIAsk ('Voulez vous sauvegarder votre saisie ?', 'Saisie DADS-U établissement') ;
if Rep=mrNo then
   exit
else
   result := TRUE;

try
   begintrans;
   DeleteErreur ('**'+Etab, 'S', 2);
   ControleOK:= ControleConform;
   if ControleOK = TRUE then
      begin
      DeleteDetail ('**'+Etab, 2);

      ChargeTOBDADS;
      SauveZones;
      LibereTOBDADS;

      TFVierge (Ecran).Binsert.Enabled:= True;
      TFVierge (Ecran).BDelete.Enabled:= True;
      end;
   CommitTrans;
Except
   result:= FALSE;
   Rollback;
   PGIBox ('Une erreur est survenue lors de la mise à jour de la base', 'Mise à jour DADS-U');
   END;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/10/2007
Modifié le ... :   /  /
Description .. : Changement assujetti
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_DADSU_ETAB.AssujettiClick(Sender: TObject);
begin
if (Tcheckbox(Sender).name='CHTASSUJETTI') then
   begin
   if (Tcheckbox(Sender).Checked=True) then
      begin
      SetControlEnabled ('LTAXE', True);
      SetControlEnabled ('ETAXE', True);
      end
   else
      begin
      SetControlEnabled ('LTAXE', False);
      SetControlEnabled ('ETAXE', False);
      SetControlText ('ETAXE', '');
      end;
   end
else
   begin
   if (Tcheckbox(Sender).Checked=True) then
      begin
//PT1
      SetControlEnabled ('LFORMATIONCDD', True);
      SetControlEnabled ('EFORMATIONCDD', True);
//FIN PT1
      SetControlEnabled ('LFORMATION', True);
      SetControlEnabled ('EFORMATION', True);
      end
   else
      begin
//PT1
      SetControlEnabled ('LFORMATIONCDD', False);
      SetControlEnabled ('EFORMATIONCDD', False);
      SetControlText ('EFORMATIONCDD', '');
//FIN PT1
      SetControlEnabled ('LFORMATION', False);
      SetControlEnabled ('EFORMATION', False);
      SetControlText ('EFORMATION', '');
      end;
   end;
end;


{$IFNDEF EAGLCLIENT}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 25/09/2007
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "SalPrem"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_DADSU_ETAB.EtabPremClick(Sender: TObject);
begin
BougeEtab (nbFirst) ;
EtabPrem.Enabled:= FALSE;
EtabPrec.Enabled:= FALSE;
EtabSuiv.Enabled:= TRUE;
EtabDern.Enabled:= TRUE;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 25/09/2007
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "SalPrec"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_DADSU_ETAB.EtabPrecClick(Sender: TObject);
begin
BougeEtab (nbPrior) ;
if QMul.BOF then
   begin
   EtabPrem.Enabled:= FALSE;
   EtabPrec.Enabled:= FALSE;
   end;
EtabSuiv.Enabled:= TRUE;
EtabDern.Enabled:= TRUE;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 25/09/2007
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "SalSuiv"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_DADSU_ETAB.EtabSuivClick(Sender: TObject);
begin
BougeEtab (nbNext) ;
EtabPrem.Enabled:= TRUE;
EtabPrec.Enabled:= TRUE;
if QMul.EOF then
   begin
   EtabSuiv.Enabled:= FALSE;
   EtabDern.Enabled:= FALSE;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 25/09/2007
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "SalDern"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_DADSU_ETAB.EtabDernClick(Sender: TObject);
begin
BougeEtab (nbLast) ;
EtabPrem.Enabled:= TRUE;
EtabPrec.Enabled:= TRUE;
EtabSuiv.Enabled:= FALSE;
EtabDern.Enabled:= FALSE;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 25/09/2007
Modifié le ... :   /  /
Description .. : Déplacement au niveau de l'établissement
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
Function TOF_DADSU_ETAB.BougeEtab(Button: TNavigateBtn) : boolean ;
BEGIN
UpdateTable;
result:= TRUE;

if Button=nbDelete then
   begin
   if QMul.EOF = FALSE then
      begin
      QMul.Next;
      if QMul.EOF = TRUE then
         begin
         QMul.prior;
         if QMul.BOF then
            Close;
         end
      end
   else
      begin
      if QMul.BOF = FALSE then
         begin
         QMul.prior;
         if QMul.BOF = TRUE then
            Close;
         end;
      end;
   end;

if QMul <> NIL then
   begin
   Case Button of
        nblast : QMul.Last;
        nbfirst : QMul.First;
        nbnext : QMul.Next;
        nbprior : QMul.prior;
        end;
   end;
GereQueryEtab;
END ;


procedure TOF_DADSU_ETAB.GereQueryEtab();
begin
if QMul = NIL then
   exit;
Etab := QMul.FindField('ET_ETABLISSEMENT').AsString;
if T_Periode <> NIL then
   FreeAndNil (T_Periode);

EtabChange:= True;
ChargeZones;
AfficheCaption;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 25/09/2007
Modifié le ... :   /  /
Description .. : Click sur le bouton Imprimer
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_DADSU_ETAB.Impression(Sender: TObject);
var
Pages : TPageControl;
Rech : String;
begin
Rech := 'SELECT *'+
        ' FROM DADSDETAIL WHERE'+
        ' PDE_SALARIE="'+Etab+'" AND'+
        ' PDE_TYPE="'+TypeD+'" AND'+
        ' PDE_EXERCICEDADS="'+PGExercice+'"';

Pages := TPageControl(GetControl('PAGES'));
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 25/09/2007
Modifié le ... :   /  /
Description .. : Click sur le bouton valider
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_DADSU_ETAB.Validation(Sender: TObject);
begin
UpdateTable;
if (ControleOK=TRUE) then
   begin
   if (State='CREATION') then
      begin
      EtabPrem.Enabled:= TRUE;
      EtabPrec.Enabled:= TRUE;
      EtabSuiv.Enabled:= TRUE;
      EtabDern.Enabled:= TRUE;
      end;
   State:= 'MODIFICATION';
   end;
end;


//PT6
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 25/09/2007
Modifié le ... :   /  /
Description .. : Complément des raccourcis claviers
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_DADSU_ETAB.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
TFVierge(Ecran).FormKeyDown(Sender, Key, Shift);
case Key of
     VK_F3: begin
            if ((GetControlVisible ('BETABPREM')) and
               (GetControlEnabled ('BETABPREM')) and (ssCtrl in Shift)) then
               EtabPrem.Click      //Premier établissement
            else
            if ((GetControlVisible ('BETABPREC')) and
               (GetControlEnabled ('BETABPREC')) and (Shift = [])) then
               EtabPrec.Click;     //Etablissement précédent
            end;
     VK_F4: begin
            if ((GetControlVisible ('BETABDERN')) and
               (GetControlEnabled ('BETABDERN')) and (ssCtrl in Shift)) then
               EtabDern.Click     //Dernier établissement
            else
            if ((GetControlVisible('BSALSUIV')) and
               (GetControlEnabled('BSALSUIV')) and (Shift = [])) then
               EtabSuiv.Click;     //Salarié suivant
            end;
     end;
end;

Initialization
registerclasses ([TOF_DADSU_ETAB]); 
end.
