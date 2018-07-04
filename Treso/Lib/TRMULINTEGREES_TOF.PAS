{ Unité : Source TOF de la FICHE : TRMULINTEGREES
  Objet : Détail d'une écriture intégrées dans la comptabilité
--------------------------------------------------------------------------------------
    Version    |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 0.91           01/08/03   JP   Création de l'unité
 7.09.001.001   07/12/06   JP   Gestion du dossier et du qualifpiece
--------------------------------------------------------------------------------------}

Unit TRMULINTEGREES_TOF ;

interface

uses
  StdCtrls, Controls,  Classes,
  {$IFDEF EAGLCLIENT}
  eMul, MaineAGL,
  {$ELSE}
  db, Mul, FE_Main,
  {$ENDIF}
  HCtrls, UTOF, HQry, SysUtils, UTOB;

type
  TOF_TRMULINTEGREES = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  private
    procedure FListeDblClick(Sender : TObject);
    procedure DeviseOnChange(Sender : TObject);
  end ;

procedure TRLanceFiche_VisuInteg(Dom, Fiche, Range, Lequel, Arguments : string);


implementation

uses
  TRLIGNEECRCOMPTA_TOF, ExtCtrls{TImage}, Commun {AssignDrapeau}, HEnt1;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_VisuInteg(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULINTEGREES.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 50000133;
  TFMul(Ecran).FListe.OnDblClick := FListeDblClick;
  THValComboBox(GetControl('E_DEVISE')).OnChange := DeviseOnChange;
  {07/12/06 : Gestion du multi dossiers}
  SetControlVisible('MULTIDOSSIER', IsTresoMultiSoc);
  SetControlVisible('TMULTIDOSSIER', IsTresoMultiSoc);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULINTEGREES.DeviseOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  AssignDrapeau(TImage(GetControl('IDEV')), GetControlText('E_DEVISE'));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULINTEGREES.FListeDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  if GetField('E_JOURNAL') = '' then Exit;
  T := TOB.Create('', nil, -1);
  try
    T.AddChampSupValeur('E_JOURNAL', GetField('E_JOURNAL'));
    T.AddChampSupValeur('E_GENERAL', GetField('E_GENERAL'));
    T.AddChampSupValeur('E_EXERCICE', GetField('E_EXERCICE'));
    T.AddChampSupValeur('E_NUMEROPIECE', Valeur(GetField('E_NUMEROPIECE')));
    T.AddChampSupValeur('E_DATECOMPTABLE', GetField('E_DATECOMPTABLE'));
    T.AddChampSupValeur('E_NUMLIGNE', GetField('E_NUMLIGNE'));
    {07/12/06 : Ajout du dossier sur laquelle effectuer la requête}
    T.AddChampSupValeur('NOMBASE', '');
    if IsTresoMultiSoc and (GetControlText('MULTIDOSSIER') <> '') then
      T.SetString('NOMBASE', GetField('SYSDOSSIER'));
    {Affichage du détail}
    AffDetailEcr(T, True);
  finally
    FreeAndNil(T);
  end;
end;


initialization
  RegisterClasses ([TOF_TRMULINTEGREES] );

end.
