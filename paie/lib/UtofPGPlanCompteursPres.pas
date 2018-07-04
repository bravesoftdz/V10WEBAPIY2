{***********UNITE*************************************************
Auteur  ...... : FLO
Créé le ...... : 04/07/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PLANCOMPTEURSPRES ()
Mots clefs ... : TOF;PLANCOMPTEURSPRES
*****************************************************************}
Unit UTOFPGPLANCOMPTEURSPRES;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     forms,
     sysutils, 
     ComCtrls,
     UTOF ;

Const NB_RUPTURES = 2;

Type
  TOF_PLANCOMPTEURSPRES = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  private
    procedure OnClickBValider(Sender : TObject);
    Function  RecupCriteresGB (GBName : String) : String;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    Procedure ChangeTheme (Sender : TObject);
    Procedure ChangeRupture (Sender : TObject);
  end ;

Implementation

Uses
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$ELSE}
     eMul,
     uTob,
{$ENDIF}
     HCtrls, 
     HEnt1,
     HMsgBox,
     HTB97, 
     P5Def,
     HQry,
     EntPaie,
     Windows,
     PgPlanningOutils,
     PGPlanningPresenceCompteurs;


procedure TOF_PLANCOMPTEURSPRES.OnArgument (S : String ) ;
Var Num,i : Integer;
    Combo : THMultiValComboBox;
    ComboRupt : THValComboBox;
begin
  Inherited ;
  
     // Action du bouton Valider
     (GetControl('BValider') as TToolBarButton97).OnClick := OnClickBValider;

     // Recherche des libellés des zones TRAVAIL, CODESTAT pour affichage ou non
     for Num := 1 to 4 do
     begin
          VisibiliteChampSalarie(IntToStr(Num), GetControl('PYP_TRAVAILN' + IntToStr(Num)), GetControl('TPYP_TRAVAILN' + IntToStr(Num)));
     end;

     VisibiliteStat(GetControl('PYP_CODESTAT'), GetControl('TPYP_CODESTAT'));

     // Gestion des raccourcis clavier
     Ecran.OnKeyDown := FormKeyDown;

     // Céation d'un listener sur le thème
     Combo := THMultiValComboBox(GetControl('PYP_THEMEPRE'));
     If Combo <> Nil then Combo.OnChange := ChangeTheme;

     // Renseignement des combos de ruptures
     For i := 1 To NB_RUPTURES Do
     Begin
          ComboRupt := GetControl('RUPTURE' + IntToStr(i)) As THValComboBox;
          If ComboRupt <> Nil Then
          Begin
                  ComboRupt.Items.Add ('<<Aucune>>');
                  ComboRupt.Values.Add('');
                  ComboRupt.Items.Add ('Etablissement');
                  ComboRupt.Values.Add('PYP_ETABLISSEMENT');
                  For Num := 1 To VH_Paie.PGNbreStatOrg Do
                  Begin
                      Case Num Of
                         1 : ComboRupt.Items.Add(VH_Paie.PGLibelleOrgStat1);
                         2 : ComboRupt.Items.Add(VH_Paie.PGLibelleOrgStat2);
                         3 : ComboRupt.Items.Add(VH_Paie.PGLibelleOrgStat3);
                         4 : ComboRupt.Items.Add(VH_Paie.PGLibelleOrgStat4);
                      End;
                      ComboRupt.Values.Add('PYP_TRAVAILN'+IntToStr(Num));
                  End;

                  ComboRupt.OnChange := ChangeRupture;
          End;
     End;
end ;


procedure TOF_PLANCOMPTEURSPRES.OnClickBValider(Sender: TObject);
var
  WhereCpt, WhereSal, Error, stwhere : String;
  NiveauRupt: TNiveauRupture;
  i : Integer;
begin
     // Contrôles de saisie
     Error := '';
     If (StrToDate(GetControlText('EDDATE')) = iDate1900) Or (StrToDate(GetControlText('EDDATE_')) = iDate1900) Then
          Error :='Vous devez choisir une période à afficher'
     Else If (StrToDate(GetControlText('EDDATE_'))-StrToDate(GetControlText('EDDATE')) > 366) Then
          Error :='Vous devez choisir une période d''affichage d''un an maximum'
     Else If (StrToDate(GetControlText('EDDATE_'))-StrToDate(GetControlText('EDDATE')) < 0) Then
          Error := 'La date de fin doit être supérieure ou égale à la date de début';

     If Error <> '' Then
     Begin
          PGIError(Error, Ecran.Caption);
          Exit;
     End;

     // Récupération des critères
     WhereCpt  := RecupCriteresGB('GBCOMPTEURS');
     WhereSal  := RecupCriteresGB('GBSALARIES');
    // Récupération des critères du code salarié
    if (getcontroltext('PYP_SALARIE') <> '') or (getcontroltext('PYP_SALARIE_') <> '') then
    begin  //au moins un code salarié de renseigné
      if (GetControlText('PYP_SALARIE') <> '') and (GetControlText('PYP_SALARIE_') <> '') then
      StWhere := StWhere + 'PYP_SALARIE>="' + GetControlText('PYP_SALARIE') + '" AND PYP_SALARIE<="' + GetControlText('PYP_SALARIE_') + '" '
      else
        if (GetControlText('PYP_SALARIE') <> '') and (GetControlText('PYP_SALARIE_') = '') then
        StWhere := StWhere + 'PYP_SALARIE>="' + GetControlText('PYP_SALARIE') + '" '
      else
        if (GetControlText('PYP_SALARIE') = '') and (GetControlText('PYP_SALARIE_') <> '') then
        StWhere := StWhere + 'PYP_SALARIE<="' + GetControlText('PYP_SALARIE_') + '" ';

      if whereSal = '' then wheresal := stwhere
      else wheresal := wheresal + ' AND ' + stwhere;
    end;

     // Récupération des ruptures
     NiveauRupt.CondRupt := '';
     NiveauRupt.NiveauRupt := 0;
     For i := 1 To NB_RUPTURES Do
     Begin
          if GetControlText('RUPTURE'+IntToStr(i)) <> '' then
          Begin
               NiveauRupt.NiveauRupt := i;
               NiveauRupt.ChampsRupt[i] := GetControlText('RUPTURE'+IntToStr(i));
          End;
     End;

     // Appel du planning
     PGPlanningPresenceCompteursOpen(StrToDate(GetControlText('EDDATE')),
                                     StrToDate(GetControlText('EDDATE_')),
                                     WhereCpt,  //TobEtats
                                     WhereSal,  //TobRessources
                                     NiveauRupt, 
                                     [piaSemaine, piaMois]);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 05/07/2007
Modifié le ... :   /  /    
Description .. : Récupération des critères d'un group box pour former la 
Suite ........ : clause Where d'une requête
Mots clefs ... : 
*****************************************************************}
function TOF_PLANCOMPTEURSPRES.RecupCriteresGB(GBName: String): String;
Var
     Where, temp, valeur : String;
     GB    : TGroupBox;
     Combo : THValComboBox;
     MCombo: THMultiValComboBox;
     Edit  : THEdit;
     i     : Integer;
begin

     Where := '';

     // Existence du group-box
     If (GBName <> '') And (GetControl(GBName) <> Nil) And (GetControl(GBName) is TGroupBox) Then
     Begin
          GB := GetControl(GBName) As TGroupBox;
          // Passage en revue des différents composants du group-box
          For i := 0 To GB.ControlCount-1 Do
          Begin
               If GB.Controls[i] Is THValComboBox Then
               Begin
                    Combo := THValComboBox (GB.Controls[i]);
                    If Combo.Value <> '' Then
                    Begin
                         If Where <> '' Then Where := Where + ' AND ';
                         Where := Where + ' ' + Combo.Name + '="' + Combo.Value + '"';
                    End;
               End
               Else If GB.Controls[i] Is THMultiValComboBox Then
               Begin
                    MCombo := THMultiValComboBox (GB.Controls[i]);
                    If MCombo.Value <> '' Then
                    Begin
                         If Where <> '' Then Where := Where + ' AND ';
                         Temp := MCombo.Value;
                         Where := Where + ' ' + MCombo.Name + ' IN (';
                         valeur := ReadTokenSt(Temp);
                         While valeur <> '' Do
                         Begin
                              Where := Where + '"' + Valeur + '"';
                              Valeur := ReadTokenSt(Temp);
                              If Valeur <> '' Then Where := Where + ',';
                         End;
                         Where := Where + ')';
                    End;
               End
               Else If GB.Controls[i] Is THEdit Then
               Begin
                    Edit := THEdit (GB.Controls[i]);
                    If Edit.Text <> '' Then
                    Begin
                         // Cas des dates
                         If Edit.DefaultDate <> odNone Then
                         Begin
                              If (Edit.Text <> '__/__/____') And (StrToDate(Edit.Text) <> iDate1900) Then
                              Begin
                                   If Where <> '' Then Where := Where + ' AND ';
                                   Where := Where + ' ' + Edit.Name;
                                   If Pos (UpperCase(Edit.Name), 'DEBUT') > 0 Then Where := Where + '>=' Else Where := Where + '<';
                                   Where := Where + '"' + UsDateTime(StrToDate(Edit.Text)) + '"';
                              End;
                         End
                         Else
                         Begin
                              If Edit.Text <> '' Then
                              Begin
                                   If Where <> '' Then Where := Where + ' AND ';
                                   Where := Where + ' ' + Edit.Name + '="' + Edit.Text + '"';
                              End;
                         End;
                    End;
               End;
          End;
     End;

     Result := Where;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 06/07/2007
Modifié le ... :   /  /    
Description .. : Gestion du raccourci clavier F9
Mots clefs ... : 
*****************************************************************}
procedure TOF_PLANCOMPTEURSPRES.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_F9:OnClickBValider (Sender);
  End;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 06/07/2007
Modifié le ... :   /  /
Description .. : Rafraîchissement de la combo des compteurs en fonction
Suite ........ : du thème choisi
Mots clefs ... : 
*****************************************************************}
procedure TOF_PLANCOMPTEURSPRES.ChangeTheme(Sender: TObject);
Var Valeur, Temp, Chaine : String;
begin
      If (THMultiValComboBox(GetControl('PYP_THEMEPRE')).Value <> '') Then
      Begin
          SetControlText('PYP_COMPTEURPRES', '');

          Chaine := '';
          Temp := GetControlText('PYP_THEMEPRE');
          Valeur := ReadTokenSt(Temp);
          While Valeur <> '' Do
          Begin
               Chaine := Chaine + '"' + Valeur + '"';
               Valeur := ReadTokenSt(Temp);
               If Valeur <> '' Then Chaine := Chaine + ',';
          End;
          SetControlProperty('PYP_COMPTEURPRES', 'Plus', ' AND PYR_THEMEPRE IN ('+Chaine+')');
      End
      Else
          SetControlProperty('PYP_COMPTEURPRES', 'Plus', '');
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/07/2007
Modifié le ... :   /  /    
Description .. : Contrôles sur validation d'une rupture
Mots clefs ... : 
*****************************************************************}
procedure TOF_PLANCOMPTEURSPRES.ChangeRupture(Sender: TObject);
Var
  i,j : Integer;
  Valeur : String;
Begin
     // Récupération du niveau de rupture
     i := StrToInt(Copy(TControl(Sender).Name,8,1));

     // Contrôle de cohérence des ruptures
     For j:= 1 To (NB_RUPTURES-1) Do
     Begin
        If (THValComboBox(GetControl('RUPTURE'+IntToStr(j))).Value = '') And
           (THValComboBox(GetControl('RUPTURE'+IntToStr(j+1))).Value <> '') Then
          Begin
               PGIBox('Le niveau de rupture '+IntToStr(j)+' doit être renseigné',Ecran.Caption);
               GetControl('BValider').Enabled := False;
               Exit;
          End
     End;

     Valeur := THValComboBox(GetControl('RUPTURE'+IntToStr(i))).Value;
     For j:=1 To NB_RUPTURES Do
     Begin
          If (i <> j) And (Valeur = THValComboBox(GetControl('RUPTURE'+IntToStr(j))).Value) And (Valeur <> '') Then
          Begin
               PGIBox('La rupture '+IntToStr(i)+' doit être différente de la '+IntToStr(j),Ecran.Caption);
               GetControl('BValider').Enabled := False;
               Exit;
          End;
     End;

     GetControl('BValider').Enabled := True;
end;

Initialization
  registerclasses ( [ TOF_PLANCOMPTEURSPRES ] ) ;
end.
