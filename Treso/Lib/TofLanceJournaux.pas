{ Unité : Source TOF de la FICHE : LANCEETAT
--------------------------------------------------------------------------------------
    Version   |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
               30/01/02  BT   Création de l'unité
 6.20.001.001  09/02/05  JP   Mise en conformité avec les OPCVM
 7.05.001.001  23/10/06  JP   Gestion des filtres multi sociétés
                              Mise en place de l'ancêtre des états
 8.00.002.002  26/07/07  JP   Nouvelle gestion des soldes
--------------------------------------------------------------------------------------}
unit TofLanceJournaux ;

Interface

uses
  StdCtrls, Controls,Classes,
  {$IFDEF EAGLCLIENT}
  MaineAGL, eQRS1 ,
  {$ELSE}
  QRS1 , FE_Main,
  {$ENDIF}
  SysUtils, HCtrls, UTOF, uAncetreEtat;


type
  TOF_LANCEJOURNAUX = class (TRANCETREETAT)
    procedure OnLoad                ; override;
    procedure OnArgument(S : string); override;
  private
    Nature	: string ;
    Modele	: string ;

    procedure NoDossierChange(Sender : TObject);
    procedure rbButtonOnClick(Sender : TObject);

    function GetDateDeb : TDateTime;
    function GetDateFin : TDateTime;
  public
    property DateDeb : TDateTime read GetDateDeb;
    property DateFin : TDateTime read GetDateFin;
  end;

procedure TRLanceFiche_LanceJourneaux(Dom, Fiche, Range, Lequel, Arguments: string);


implementation

uses
  HEnt1, Commun, UProcGen;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_LanceJourneaux(Dom, Fiche, Range, Lequel, Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_LANCEJOURNAUX.OnLoad ;
{---------------------------------------------------------------------------------------}
var
  Nat : string;
  DtA : string;
begin
  inherited;
  Nat := '';
  DtA := DateToStr(DebutAnnee(V_PGI.DateEntree));

  if (GetControlText('TE_NATURE') <> '') and (Pos('<<', GetControlText('TE_NATURE')) = 0) then begin
    Nat := 'AND TE_NATURE IN(' + GetClauseIn(GetControlText('TE_NATURE')) + ')';
  end;
  SetControlText('NATURE', Nat);

  if DateDeb > iDate1900 then
    DtA := DateToStr(DebutAnnee(DateDeb));
  SetControlText('DEBANNEE', DtA);
  if DtA = DateToStr(DateDeb) then SetControlText('DEBSQL', DateToStr(StrToDate(DtA) + 1))
                              else SetControlText('DEBSQL', DateToStr(DateDeb));

  SetControlText('PERDEB', DateToStr(DateDeb));
  SetControlText('PERFIN', DateToStr(DateFin));
  
  if (GetControl('RBDATEVAL') as TRadioButton).Checked then begin
    SetControlText('LIBDATE', TraduireMemoire('Date de valeur'));
    SetControlText('CHAMP', 'TE_DATEVALEUR');
  end
  else begin
    SetControlText('LIBDATE', TraduireMemoire('Date d''opération'));
    SetControlText('CHAMP', 'TE_DATECOMPTABLE');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_LANCEJOURNAUX.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited ;

  Nature := ReadTokenSt(S);
  Modele := ReadTokenSt(S);

  TFQRS1(Ecran).HelpContext := 150;
  TFQRS1(Ecran).CodeEtat:=Modele;
  TFQRS1(Ecran).NatureEtat:=Nature;

  {08/08/06 : gestion du multi sociétés}
  SetControlVisible('TE_NODOSSIER' , IsTresoMultiSoc);
  SetControlVisible('TTE_NODOSSIER', IsTresoMultiSoc);

  {Gestion des filtres multi sociétés sur banquecp et dossier}
  if not EtatMD then begin
    THEdit(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(THEdit(GetControl('TE_GENERAL')).DataType, '', '');
    THEdit(GetControl('TE_GENERAL_')).Plus := FiltreBanqueCp(THEdit(GetControl('TE_GENERAL_')).DataType, '', '');
    THValComboBox(GetControl('TE_NODOSSIER')).Plus := 'DOS_NODOSSIER ' + FiltreNodossier;
    THValComboBox(GetControl('TE_NODOSSIER')).OnChange := NoDossierChange;
  end;

  (GetControl('RBDATEVAL') as TRadioButton).OnClick := rbButtonOnClick;
  (GetControl('RBDATEOPE') as TRadioButton).OnClick := rbButtonOnClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_LANCEJOURNAUX.NoDossierChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  THEdit(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(THEdit(GetControl('TE_GENERAL')).DataType, '', GetControlText('TE_NODOSSIER'));
  SetControlText('TE_GENERAL', '');
  THEdit(GetControl('TE_GENERAL_')).Plus := FiltreBanqueCp(THEdit(GetControl('TE_GENERAL_')).DataType, '', GetControlText('TE_NODOSSIER'));
  SetControlText('TE_GENERAL_', '');
end;

{---------------------------------------------------------------------------------------}
function TOF_LANCEJOURNAUX.GetDateDeb : TDateTime;
{---------------------------------------------------------------------------------------}
begin
  if (GetControl('RBDATEVAL') as TRadioButton).Checked then begin
    if IsValidDate(GetControlText('TE_DATEVALEUR')) then Result := StrToDate(GetControlText('TE_DATEVALEUR'))
                                                    else Result := DebutAnnee(V_PGI.DateEntree);
  end
  else begin
    if IsValidDate(GetControlText('TE_DATECOMPTABLE')) then Result := StrToDate(GetControlText('TE_DATECOMPTABLE'))
                                                       else Result := DebutAnnee(V_PGI.DateEntree);
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_LANCEJOURNAUX.GetDateFin : TDateTime;
{---------------------------------------------------------------------------------------}
begin
  if (GetControl('RBDATEVAL') as TRadioButton).Checked then begin
    if IsValidDate(GetControlText('TE_DATEVALEUR_')) then Result := StrToDate(GetControlText('TE_DATEVALEUR_'))
                                                     else Result := DebutAnnee(V_PGI.DateEntree);
  end
  else begin
    if IsValidDate(GetControlText('TE_DATECOMPTABLE_')) then Result := StrToDate(GetControlText('TE_DATECOMPTABLE_'))
                                                        else Result := DebutAnnee(V_PGI.DateEntree);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_LANCEJOURNAUX.rbButtonOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Bok : Boolean;
begin
  Bok := (GetControl('RBDATEVAL') as TRadioButton).Checked;
  SetControlVisible('TTE_DATEVALEUR', Bok);
  SetControlVisible('TE_DATEVALEUR' , Bok);
  SetControlVisible('TE_DATEVALEUR_', Bok);
  SetControlVisible('TTE_DATECOMPTABLE', not Bok);
  SetControlVisible('TE_DATECOMPTABLE' , not Bok);
  SetControlVisible('TE_DATECOMPTABLE_', not Bok);
end;

Initialization
  registerclasses ( [ TOF_LANCEJOURNAUX ] ) ;
end.

