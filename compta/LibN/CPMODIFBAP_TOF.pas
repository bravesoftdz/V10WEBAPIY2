{ Unité : Source TOF de la fiche CPMODIFBAP
--------------------------------------------------------------------------------------
    Version    |   Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 7.01.001.001   20/02/06   JP   Création de l'unité
 8.01.001.001   24/01/07   JP   FQ 19587 : ID HelpTopic
 8.01.001.001   24/01/07   JP   FQ 19448 : Ajout bouton accès aux circuits
 8.01.001.012   24/04/07   JP   Modification de la requête de mise à jour sur BAP_ECHEANCEBAP
 8.00.001.018   29/05/07   JP   Mise en place des rôles à la place des groupes pour filtrer les utilisateurs
--------------------------------------------------------------------------------------}
unit CPMODIFBAP_TOF;

interface

uses
  Controls, Classes, Vierge,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  Grids, SysUtils, HCtrls, UTob, HEnt1, uLibBonAPayer, UTOF;

type
  TOF_CPMODIFBAP = class(TOF)
    FListe   : THGrid;

    procedure OnArgument(S : string); override;
    procedure OnUpdate              ; override;
  private
    TobTraitement : TOB;
    ObjCircuit : TObjCircuit;

    function  PrepareRequete : string;
    procedure AccesCircuit(Sender : TObject); {JP 24/01/07 : FQ 10448}
  public
    procedure MajControl  (Sender : TObject);
    procedure ElipsisClick(Sender : TObject); {29/05/07}
  end;

procedure CpLanceFiche_ModifBap;

implementation

uses
  HMsgBox, Forms, HTB97, CPCIRCUITBAP_TOF {FQ 10448};

{---------------------------------------------------------------------------------------}
procedure CpLanceFiche_ModifBap;
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('CP', 'CPMODIFBAP', '', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFBAP.MajControl(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {On ne peut modifier le circuit et les viseurs}
  if TWinControl(Sender).Name = 'CBCIRCUIT' then begin
    SetControlEnabled('CBVISEUR1', GetControlText('CBCIRCUIT') = '');
    SetControlEnabled('CBVISEUR2', GetControlText('CBCIRCUIT') = '');
    SetControlEnabled('LBVISEUR1', GetControlText('CBCIRCUIT') = '');
    SetControlEnabled('LBVISEUR2', GetControlText('CBCIRCUIT') = '');
    SetControlText('CBVISEUR1', '');
    SetControlText('CBVISEUR2', '');
  end
  else begin
    SetControlEnabled('CBCIRCUIT', (GetControlText('CBVISEUR1') = '') and (GetControlText('CBVISEUR2') = ''));
    SetControlEnabled('LBCIRCUIT', (GetControlText('CBVISEUR1') = '') and (GetControlText('CBVISEUR2') = ''));
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFBAP.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 7509070; {JP 24/01/07 : FQ 19587}

  TobTraitement := TFVierge(Ecran).LaTOF.LaTOB;

  {29/05/07 : Mise en place des rôles à la place des groupes pour filtrer les utilisateurs}
  if GetControl('CBVISEUR1') is THEdit then begin
    (GetControl('CBVISEUR1') as THEdit).OnElipsisClick := ElipsisClick;
    (GetControl('CBVISEUR2') as THEdit).OnElipsisClick := ElipsisClick;
    (GetControl('CBVISEUR1') as THEdit).OnChange := MajControl;
    (GetControl('CBVISEUR2') as THEdit).OnChange := MajControl;
  end else begin
    SetPlusCombo(THValComboBox(GetControl('CBVISEUR1')), 'US');
    SetPlusCombo(THValComboBox(GetControl('CBVISEUR2')), 'US');
    THValComboBox(GetControl('CBVISEUR1')).OnChange := MajControl;
    THValComboBox(GetControl('CBVISEUR2')).OnChange := MajControl;
  end;

  THValComboBox(GetControl('CBCIRCUIT')).OnChange := MajControl;
  (GetControl('BCIRCUIT') as TToolbarButton97).OnClick := AccesCircuit;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFBAP.OnUpdate;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  s : string;
  w : string;
  c : string;
  T : TOB;
  V1 : string;
  V2 : string;
  Nb : Integer;
  Jo : Integer;
  Ok : Boolean;
begin
  inherited;
  if HShowMessage('0;' + Ecran.Caption + ';Êtes vous sûr de vouloir modifier les bons à payer sélectionnés ?;Q;YN;N;N;', '', '') <> mrYes then
    Exit;

  if not Assigned(TobTraitement) then Exit;

  if GetControlText('CBCIRCUIT') <> '' then
    ObjCircuit := TObjCircuit.Create(GetControlText('CBCIRCUIT'));
  try
    Ok := False;
    c := PrepareRequete;
    for n := 0 to TobTraitement.Detail.Count - 1 do begin
      T := TobTraitement.Detail[n];
      w := GetWhereBapT(T);
      if Assigned(ObjCircuit) then begin
        V1 := ObjCircuit.GetViseur1(T.GetInteger('BAP_NUMEROORDRE'));
        V2 := ObjCircuit.GetViseur2(T.GetInteger('BAP_NUMEROORDRE'));
        Nb := ObjCircuit.GetDelais(T.GetInteger('BAP_NUMEROORDRE'));
        {Il se peut que le précédent circuit ait eu plus de lignes que le nouveau et
         que la ligne BAP que l'on cherche à modifier soit supérieure au nombre de
         lignes du nouveau circuit}
        if (Nb = -1) or (V1 = '') or (V2 = '') then begin
          Ok := True;
          Continue;
        end
        else begin
          s := c + ', BAP_VISEUR1 = "' + V1 + '" ';
          s := s + ', BAP_VISEUR2 = "' + V2 + '" ';
          {24/04/07 : Pour éviter qu'il y ait deux maj sur BAP_ECHEANCEBAP !!}
          if IsValidDate(GetControlText('EDDATE')) and (StrToDate(GetControlText('EDDATE')) > iDate1900) then begin
            s := s + ', BAP_ECHEANCEBAP = "' + USDateTime(StrToDate(GetControlText('EDDATE'))) + '" ';
          end
          else begin
            {Pour calculer la nouvelle date d'échéance du BAP, on va partir de la différence entre
             l'ancien et le nouveau délais. Si le NbJour est à 0, c'est qu'il y a urgence => on ne
             touche à rien}
            if T.GetInteger('BAP_NBJOUR') > 0 then Jo := Nb - T.GetInteger('BAP_NBJOUR')
                                              else Jo := 0;
            s := s + ', BAP_NBJOUR = ' + IntToStr(Nb) + ' ';
            s := s + ', BAP_ECHEANCEBAP = BAP_ECHEANCEBAP + ' + IntToStr(Jo) + ' ';
          end;
        end
      end
      {On ne modifie pas le circuit, mais un des autres paramètres}
      else begin
        {24/04/07 : Déplacer depuis le PrepareRequete}
        if IsValidDate(GetControlText('EDDATE')) and (StrToDate(GetControlText('EDDATE')) > iDate1900) then
          s := c + ', BAP_ECHEANCEBAP = "' + USDateTime(StrToDate(GetControlText('EDDATE'))) + '" ';
        s := c;
      end;
        
      ExecuteSQL(s  + w);

    end;
    if Ok then
      HShowMessage('0;' + Ecran.Caption + ';Certaines lignes n''ont pu être modifiées car leur indice'#13 +
                   'était supérieur au nombre de viseurs du circuit.;W;O;O;O;', '', '');
  finally
    if Assigned(ObjCircuit) then FreeAndNil(ObjCircuit);
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPMODIFBAP.PrepareRequete : string;
{---------------------------------------------------------------------------------------}
begin
  Result := 'UPDATE CPBONSAPAYER SET BAP_MODIFICATEUR = "' + V_PGI.User + '", BAP_DATEMODIF = "' + UsDateTime(Now) + '", ' +
            'BAP_STATUTBAP = "' + sbap_Encours + '" '; {Si le bap est refusé, on le remet à En cours}
  {JP 24/04/07 : Déplacer dans le OnUpdate
  if IsValidDate(GetControlText('EDDATE')) and (StrToDate(GetControlText('EDDATE')) > iDate1900) then
    Result := Result + ', BAP_ECHEANCEBAP = "' + USDateTime(StrToDate(GetControlText('EDDATE'))) + '" ';}
  if GetControlText('CBVISEUR1') <> '' then
    Result := Result + ', BAP_VISEUR1 = "' + GetControlText('CBVISEUR1') + '" ';
  if GetControlText('CBVISEUR2') <> '' then
    Result := Result + ', BAP_VISEUR2 = "' + GetControlText('CBVISEUR2') + '" ';
  if GetControlText('CBCIRCUIT') <> '' then
    Result := Result + ', BAP_CIRCUITBAP = "' + GetControlText('CBCIRCUIT') + '" ';
end;

{JP 24/01/07 : FQ 10448 : accès aux circuits de validations
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFBAP.AccesCircuit(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if GetControlText('CBCIRCUIT') = '' then
    HShowMessage('0;' + Ecran.Caption + ';Veuillez choisir un circuit.;W;O;O;O;', '', '')
  else
    CpLanceFiche_CircuitFiche(GetControlText('CBCIRCUIT'), 'M');
end;

{JP 29/05/07 : Gestion des rôles et passage en lookup des combos
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFBAP.ElipsisClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  LookUpUtilisateur(Sender);
end;

initialization
  RegisterClasses([TOF_CPMODIFBAP]);

end.
