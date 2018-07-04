{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/10/2006
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : PAIE, PGDADSU
*****************************************************************}
{
PT1    : 09/11/2007 VG V_80 Prise en compte des FQ N°13646 et 14035 
}
unit UTofPG_MULDADSUAFFECT;

interface
uses
{$IFNDEF EAGLCLIENT}
  db,
  dbTables,
  HDB,
  DBGrids,
  mul,
  FE_Main,
{$ELSE}
  MaineAgl,
  eMul,
{$ENDIF}
  StdCtrls,
  Controls,
  Classes,
  Graphics,
  forms,
  sysutils,
  ComCtrls,
  HTB97,
  Grids,
  HCtrls,
  HEnt1,
  vierge,
  EntPaie,
  HMsgBox,
  Hqry,
  UTOF,
  UTOB,
  UTOM,
  AGLInit,
  ed_tools,
  hstatus,
  PgOutils,
  PgOutils2;
  
type
  TOF_PGMUL_DADSUAFFECT = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  private
    Q_Mul: THQuery;
    BCherche, Delete : TToolbarButton97;
    CEG, DOS, STD : boolean;
    procedure ActiveWhere;
    procedure GrilleDblClick(Sender: TObject);
    procedure DeleteClick (Sender: TObject);
    procedure Delete_un();
  end;

implementation

//uses P5Def;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/10/2006
Modifié le ... :   /  /
Description .. : OnArgument
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMUL_DADSUAFFECT.OnArgument(Arguments: string);
var
{$IFDEF EAGLCLIENT}
Grille : THGrid;
{$ELSE}
Grille : THDBGrid;
{$ENDIF}
begin
inherited;
SetControlText ('PUO_NATURERUB', 'BAS');
BCherche:= TToolbarButton97 (GetControl ('BCherche'));
Q_Mul:= THQuery (Ecran.FindComponent ('Q'));
{$IFDEF EAGLCLIENT}
Grille:= THGrid (GetControl ('Fliste'));
{$ELSE}
Grille:= THDBGrid (GetControl ('Fliste'));
{$ENDIF}
if (Grille <> nil) then
   Grille.OnDblClick:= GrilleDblClick;

Delete:= TToolbarButton97 (GetControl ('BDELETE'));
if (Delete <> nil) then
   begin
   Delete.Visible:= True;
   Delete.Enabled:= True;
   Delete.OnClick:= DeleteClick;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/10/2006
Modifié le ... :   /  /
Description .. : OnLoad
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMUL_DADSUAFFECT.OnLoad;
begin
inherited;
ActiveWhere;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/10/2006
Modifié le ... :   /  /
Description .. : ActiveWhere
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMUL_DADSUAFFECT.ActiveWhere;
var
WW: THEdit;
begin
WW:= THEdit (GetControl ('XX_WHERE'));
if (WW <> nil) then
   begin
   if ((Q_Mul <> nil) and (GetControlText('PUO_NATURERUB')='AAA')) then
      TFMul(Ecran).SetDBListe('PGDADSUAFFECTREM')
   else
      TFMul(Ecran).SetDBListe('PGDADSUAFFECTCOT');
   WW.text:= '';
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/10/2006
Modifié le ... :   /  /
Description .. : GrilleDblClick
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMUL_DADSUAFFECT.GrilleDblClick(Sender: TObject);
var
NatureRub, Rubrique, Segment : string;
begin
if ((Q_Mul <> nil) and (Q_Mul.RecordCount = 0)) then
   exit;

{$IFDEF EAGLCLIENT}
Q_Mul.TQ.Seek (TFmul (Ecran).Fliste.Row-1);
{$ENDIF}

NatureRub:= Q_Mul.FindField ('PUO_NATURERUB').AsString;
Rubrique:= Q_Mul.FindField ('PUO_RUBRIQUE').AsString;
Segment:= Q_Mul.FindField ('PUO_UTILSEGMENT').asstring;//PT1
AglLanceFiche ('PAY', 'DADSUAFFECT', '', NatureRub+';'+Rubrique+';'+Segment,
               NatureRub);
if (BCherche <> nil) then
   BCherche.click;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 09/10/2006
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "Delete"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PGMUL_DADSUAFFECT.DeleteClick(Sender: TObject);
Var
i, reponse : integer;
{$IFNDEF EAGLCLIENT}
Liste : THDBGrid;
{$ELSE}
Liste : THGrid;
{$ENDIF}
begin
{$IFNDEF EAGLCLIENT}
Liste:= THDBGrid (GetControl ('FListe'));
{$ELSE}
Liste:= THGrid (GetControl ('FListe'));
{$ENDIF}

reponse:= PGIAsk ('Cette commande supprimera les affectations sélectionnées.#13#10'+
                  'Voulez-vous continuer ?', TFMul (Ecran).Caption);
if (reponse <> mrYes) then
   exit;

if Liste <> nil then
   begin
   if (Liste.NbSelected = 0) and (not Liste.AllSelected) then
      begin
      MessageAlerte ('Aucun élément sélectionné');
      exit;
      end;

   if (Liste.AllSelected = TRUE) then
      begin
      InitMoveProgressForm (nil, 'Suppression en cours',
                            'Veuillez patienter SVP ...',
                            TFmul (Ecran).Q.RecordCount, FALSE, TRUE);
      InitMove (TFmul (Ecran).Q.RecordCount, '');

{$IFDEF EAGLCLIENT}
      if (TFMul (Ecran).bSelectAll.Down) then
         TFMul (Ecran).Fetchlestous;
{$ENDIF}
      TFmul (Ecran).Q.First;
      while not TFmul (Ecran).Q.EOF do
            begin
            Delete_un;
            TFmul (Ecran).Q.Next;
            end;

      Liste.AllSelected:= False;
      TFMul (Ecran).bSelectAll.Down:= Liste.AllSelected;
      end
   else
      begin
      InitMoveProgressForm (NIL,'Suppression en cours',
                            'Veuillez patienter SVP ...',
                            Liste.NbSelected, FALSE, TRUE);
      InitMove(Liste.NbSelected, '');

      for i:= 0 to Liste.NbSelected-1 do
          begin
          Liste.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
          TFMul (Ecran).Q.TQ.Seek (TFMul (Ecran).FListe.Row-1) ;
{$ENDIF}
          Delete_un;
          end;

      Liste.ClearSelected;
      end;

   FiniMove;
   FiniMoveProgressForm;
   PGIBox ('Traitement terminé', TFMul (Ecran).Caption);
   end;

if BCherche <> nil then
   BCherche.click;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 09/10/2006
Modifié le ... :   /  /
Description .. : Suppression d'une seule affectation
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMUL_DADSUAFFECT.Delete_un();
var
Nature, NoDossier, Predefini, Rubrique, Segment, St : String;
DeleteOK : Boolean;
begin
Predefini:= TFmul (Ecran).Q.FindField ('PUO_PREDEFINI').asstring;//PT1
Nature:= TFmul (Ecran).Q.FindField ('PUO_NATURERUB').asstring;
Rubrique:= TFmul (Ecran).Q.FindField ('PUO_RUBRIQUE').asstring;
Segment:= TFmul (Ecran).Q.FindField ('PUO_UTILSEGMENT').asstring;//PT1
try
   begintrans;
//PT1
   AccesPredefini ('TOUS', CEG, STD, DOS);
   if (Predefini='CEG') then
      begin
      DeleteOK:= CEG;
      NoDossier:= '000000';
      end
   else
   if (Predefini='STD') then
      begin
      DeleteOK:= STD;
      NoDossier:= '000000';
      end
   else
   if (Predefini='DOS') then
      begin
      DeleteOK:= DOS;
      NoDossier:= PgRendNoDossier;
      end;

   if (DeleteOK) then
      begin
      St:= 'DELETE FROM PUBLICOTIS WHERE ##PUO_PREDEFINI##'+
           ' PUO_NATURERUB="'+Nature+'" AND'+
           ' PUO_RUBRIQUE="'+Rubrique+'" AND'+
           ' PUO_UTILSEGMENT="'+Segment+'"';
//FIN PT1
      ExecuteSQL (St);
      end;
   CommitTrans;
except
   Rollback;
   end;
MoveCur(False);
MoveCurProgressForm (St);
end;


initialization
  registerclasses([TOF_PGMUL_DADSUAFFECT]);
end.

