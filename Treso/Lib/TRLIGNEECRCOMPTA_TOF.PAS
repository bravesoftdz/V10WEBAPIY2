{ Unité : Source TOF de la FICHE : TRLIGNEECRCOMPTA
  Objet : Détail d'une écriture à intégrer dans la comptabilité
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
 0.91            28/07/03   JP   Création de l'unité
 0.91            27/08/03   JP   Gestion des écritures intégrées en compta
 7.09.001.001    07/12/06   JP   Gestion du dossier et du qualifpiece
--------------------------------------------------------------------------------------}
unit TRLIGNEECRCOMPTA_TOF ;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  Controls, Classes, Windows, Messages,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
  {$ENDIF}
  Forms, SysUtils, HCtrls, HEnt1, HMsgBox, UTOF, UTOB, ExtCtrls;

type
  TOF_TRLIGNEECRCOMPTA = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    IntegreeOk : Boolean;
    T          : Tob;
  end ;

procedure AffDetailEcr(TobLigne : TOB; Integree : Boolean = False);

implementation

uses
  Commun, UtilPGI;

var
  TobDetail : TOB;

{---------------------------------------------------------------------------------------}
procedure AffDetailEcr(TobLigne : TOB; Integree : Boolean = False);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  if not Assigned(TobLigne) then Exit;
  TobDetail := TobLigne;
  if Integree then s := 'OK;';
  AGLLanceFiche('TR', 'TRLIGNEECRCOMPTA', '', '', s);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRLIGNEECRCOMPTA.OnLoad ;
{---------------------------------------------------------------------------------------}
var
  Dev : string;
  n   : Byte;
  C   : TComponent;
begin
  inherited;
  {Mise à jour du libellé de l'écran}
  if IntegreeOk then
    Ecran.Caption := 'Écriture comptable : ' + GetControlText('E_LIBELLE')
  else
    Ecran.Caption := 'Écriture comptable à intégrer : ' + GetControlText('E_LIBELLE');

  UpdateCaption(Ecran);

  {Mise en lecture seule de tous les THEdit}
  for n := 0 to Ecran.ComponentCount - 1 do begin
    Ecran.KeyPreview := False;
    C := Ecran.Components[n];
    if C is THEdit then begin
      THEdit(C).Enabled := False;
      THEdit(C).OnDblClick := nil;
    end;
  end;

  {Affichage du drapeau devise et du libellé de la devise de travail}
  Dev := GetControlText('E_DEVISE');
  AssignDrapeau(TImage(GetControl('IDEV')), Dev);
  SetControlText('E_DEVISE', RechDom('TTDEVISE', Dev, False));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRLIGNEECRCOMPTA.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  Dat : string;
  Q   : TQuery;
begin
  inherited;
  Ecran.HelpContext := 150;
  IntegreeOk := ReadTokenSt(S) = 'OK';
  {JP 27/08/03 : S'il s'agit des écritures déjà intégrées, on crée la tob}
  if IntegreeOk then begin
    Dat := TobDetail.GetValue('E_DATECOMPTABLE');
    T := Tob.Create('$Divers', nil, -1);
    SQL := 'SELECT E_JOURNAL, E_GENERAL, E_REFINTERNE, E_DEBIT, E_CREDIT, E_NUMEROPIECE,' +
           ' E_NUMLIGNE, E_MODEPAIE, E_LIBELLE, E_DATECOMPTABLE, E_EXERCICE, E_NATUREPIECE,' +
           ' E_QUALIFORIGINE, E_DEVISE, E_DATECREATION, E_DATEECHEANCE' +
           {07/12/06 : Ajout du dossier et du qualifpiece}
           ' FROM ' + GetTableDossier(TobDetail.GetString('NOMBASE'), 'ECRITURE') + ' WHERE E_QUALIFPIECE = "N" AND' +
           ' E_JOURNAL     = "' + TobDetail.GetValue('E_JOURNAL')     + '" AND' +
           ' E_GENERAL     = "' + TobDetail.GetValue('E_GENERAL')     + '" AND' +
           ' E_EXERCICE    = "' + TobDetail.GetValue('E_EXERCICE')    + '" AND' +
           ' E_NUMEROPIECE =  ' + VartoStr(TobDetail.GetValue('E_NUMEROPIECE')) + ' AND' +
           ' E_DATECOMPTABLE ="' + USDateTime(StrToDate(Dat)) + '" AND' + //
           ' E_NUMLIGNE     = ' + VartoStr(TobDetail.GetValue('E_NUMLIGNE'));
    Q := OpenSQL(SQL, True);
    if Q.RecordCount > 0 then begin
      T.LoadDetailDB('', '', '', Q, False);
      T.Detail[0].PutEcran(Ecran);
    end
    else begin
      MessageAlerte(TraduireMemoire('Impossible de récupérer l''enregistrement en cours'));
      SendMessage(Ecran.Handle, WM_CLOSE, 0, 0);
      Ferme(Q);
    end;
  end
  else
    {Chargement de Tob fille}
    TobDetail.PutEcran(Ecran);

end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRLIGNEECRCOMPTA.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(T) then FreeAndNil(T);
  inherited;
end;

initialization
  RegisterClasses ( [ TOF_TRLIGNEECRCOMPTA ] ) ;

end.
