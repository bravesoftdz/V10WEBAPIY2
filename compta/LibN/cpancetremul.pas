{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
08.10.002.001  25/10/07   JP   FQ 19970 : on limite la gestion de l'établissement par défaut
--------------------------------------------------------------------------------------}

unit CPANCETREMUL;

interface

uses
  UTOF, Controls;

type
  TOFANCETREMUL = class (TOF)
    procedure OnArgument (S : string); override;
  private
    FComboEtab : TControl;
    function  GetComboEtab : TControl; {JP 28/06/06 : FQ 16149}
  public
    {JP 15/10/07 : FQ 16149 : gestion des réstrictions Etablissements et à défaut des ParamSoc}
    procedure GereEtablissement;
    procedure MyAfterSelectFiltre; virtual;
    procedure ControlEtab;

    property  ComboEtab : TControl read GetComboEtab write FComboEtab; {JP 28/06/06 : FQ 16149}
  end;

implementation

uses
  {$IFDEF EAGLCLIENT}
  eMul,
  {$ELSE}
  Mul,
  {$ENDIF EAGLCLIENT}
  Ent1, HCtrls;

{---------------------------------------------------------------------------------------}
procedure TOFANCETREMUL.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  FComboEtab := nil;
  {JP 15/10/07 : FQ 16149 : gestion des réstrictions Etablissements et à défaut des ParamSoc}
  GereEtablissement;
  {JP 15/10/07 : FQ 16149 : on s'assure que le filtre ne va pas à l'encontre des restrictions utlisateurs}
  TFMul(Ecran).OnAfterSelectFiltre := MyAfterSelectFiltre;
end;

{JP 15/10/07 : FQ 16149 : gestion des réstrictions Etablissements et à défaut des ParamSoc
{---------------------------------------------------------------------------------------}
procedure TOFANCETREMUL.GereEtablissement;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(ComboEtab) then begin
    {Si l'on ne gère pas les établissement ...}
    if not VH^.EtablisCpta  then begin
      {... on affiche l'établissement par défaut}
      SetControlText(ComboEtab.Name, VH^.EtablisDefaut);
      {... on désactive la zone}
      SetControlEnabled(ComboEtab.Name, False);
    end

    {On gère l'établissement, donc ...}
    else begin
      {... On commence par regarder les restrictions utilisateur}
      PositionneEtabUser(ComboEtab);
      {... s'il n'y a pas de restrictions, on reprend le paramSoc
       JP 25/10/07 : FQ 19970 : Finalement on oublie l'option de l'établissement par défaut
      if GetControlText(ComboEtab.Name) = '' then begin
        {... on affiche l'établissement par défaut
        SetControlText(ComboEtab.Name, VH^.EtablisDefaut);
        {... on active la zone
        SetControlEnabled(ComboEtab.Name, True);
      end;}
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOFANCETREMUL.MyAfterSelectFiltre;
{---------------------------------------------------------------------------------------}
begin
  {JP 15/10/07 : FQ 16149 : on s'assure que le filtre ne va pas à l'encontre des
               restrictions utlisateurs}
  ControlEtab;
end;

{JP 15/10/07 : FQ 16149 : on s'assure que le filtre ne va pas à l'encontre des
               restrictions utlisateurs
{---------------------------------------------------------------------------------------}
procedure TOFANCETREMUL.ControlEtab;
{---------------------------------------------------------------------------------------}
var
  Eta : string;
begin
  if not Assigned(ComboEtab) then Exit;
  {S'il n'y a pas de gestion des établissement, logiquement, on ne force pas l'établissement !!!}
  if not VH^.EtablisCpta then Exit;

  Eta := EtabForce;
  {S'il y a une restriction utilisateur et qu'elle ne correspond pas au contenu de la combo ...}
  if (Eta <> '') and (Eta <> GetControlText(ComboEtab.Name)) then begin
    {... on affiche l'établissement des restrictions}
    SetControlText(ComboEtab.Name, Eta);
    {... on désactive la zone}
    SetControlEnabled(ComboEtab.Name, False);
  end;
end;

{Récupération du combo gérant les établissements
{---------------------------------------------------------------------------------------}
function TOFANCETREMUL.GetComboEtab : TControl;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  if FComboEtab = nil then begin
    for n := 0 to Ecran.ComponentCount - 1 do
      if (Pos('ETAB', Ecran.Components[n].Name) > 0) and
         ((Ecran.Components[n] is THValComboBox) or
          (Ecran.Components[n] is THMultiValComboBox)) then begin
        FComboEtab := TControl(Ecran.Components[n]);
        Break;
      end;
  end;
  Result := FComboEtab;
end;

end.
