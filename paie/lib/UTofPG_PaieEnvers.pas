{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 04/09/2001
Modifié le ... :   /  /
Description .. : Utof de calcul de la paie à l'envers
Suite ........ : Le principe consiste à calculer le net à payer du bulletin en
Suite ........ : partant d'un valeur approchée du brut et de travailler par
Suite ........ : dichotomie.
Mots clefs ... : PAIE;PGBULLETIN
*****************************************************************}
{
PT1   : 16/01/2003 PH V591 Blocage calcul si pas de rem saisissable
PT2   : 26/09/2003 PH V_421 FQ 10795 Si calcul n'a pas abaouti alors message indicatif
PT3   : 10/03/2004 PH V_500 Publication de la fonction CalculPEnvers
PT4   : 28/06/2004 PH V_500 Nouvelle méthode de calcul : à partir du brut Cumul 01
PT5   : 29/06/2004 PH V_500 Nouvelle méthode de calcul : à partir du brut Habituel Cumul 05
PT6   : 09/08/2004 PH V_500 FQ 11471 Blocage du lancement du calcul si anomalie détectée
PT7   : 17/05/2005 PH V_600 FQ 11801 gestion arrondi calcul paie à l'envers
PT8   : 02/08/2005 PH V_600 FQ 14484 Affichage de la méthode retenue pour le calcul
------- 03/03/2005 PH suppression des conseils et avertissements ------------------------
PT9   : 28/02/2006 MF V_65  FQ 12827 Exclusion de rub  de rem
PT10  : 27/09/2006 MF V_70  FQ 13528 Correction Exclusion de rub de rem de la paie à l'envers
PT11  : 03/05/2007 MF V_72  FQ 14088 Appel du calcul des IJSS sur clique droit
PT12  : 05/06/2007 MF V_72  FQ 14338 Ajout sablier après click sur CALC
PT13-1: 09/08/2007 PH V_72  FQ 14623 1) Permettre la paie à l'envers à partir d'une rubrique en net saisie de bulletin

}
unit UTofPG_PaieEnvers;

interface
uses Windows,
  DB,
  StdCtrls,
  Controls,
  Classes,
  Graphics,
  forms,
  sysutils,
  ComCtrls,
  HTB97,
{$IFDEF EAGLCLIENT}
  UtileAGL,
{$ELSE}
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  uPaieRemunerations,
  Grids,
  HCtrls,
  HEnt1,
  EntPaie,
  HMsgBox,
{$IFNDEF EAGLSERVER}
  UTOF,
  Vierge,
{$ENDIF}
  UTOB,
  P5Util,
  P5Def,
  SaisBul,
  PgCongesPayes,
  Buttons, // PT11
//  AGLInit,
  ed_tools;
//  HQuickRP,

// PT7 Rajout paramètre pour gestion de l'arrondi
// PT9 Rajout paramètre pour exclusion de rub de rem
function CalculPEnvers(var PlusApprochant: Double; TOB_Lignes, TLig: TOB; NetTrouve: Double; Rub: string; Silencieux: Boolean; GestArrondi: Boolean = FALSE; CalcMaint: Boolean = FALSE): Boolean;

{$IFNDEF EAGLSERVER}
type
  TOF_PG_PaieEnvers = class(TOF)
  private
    TLig, TOB_Lignes: TOB; // Tob correspondant à la ligne sur lequel on cherche le montant pour faire le calcul de la paie à l'envers
    RUBDUBRUT: THEdit;
    NETAPAYER: THNumEdit;
// d PT11
    MtIJNettes: THNumEdit;
    MtIJBrutes: THNumEdit;
    MtGarantieNet: THNumEdit;
    RIJNettes, RIJBrutes, RGarantieNet: string;
// f PT11
    RubNet: Boolean; // PT13-1
    procedure RubChange(Sender: TObject);
    function RendRubPEnvers(TOB_Lignes: TOB): Boolean;
    procedure BENVERSClick(Sender: TObject);
// d PT11
    procedure RechRubIJSS(var RIJNettes, LIJNettes, RIJBrutes, LIJBrutes, RGarantieNet, LGarantieNet: string);
    procedure CalcRubIJSS(Sender: TObject);
    procedure BValIJClick(Sender: TObject);
// f PT11
  public
    procedure OnArgument(Arguments: string); override;
  end;

{$ENDIF}
implementation

uses
  P5RecupInfos;

{$IFNDEF EAGLSERVER}

procedure TOF_PG_PaieEnvers.OnArgument(Arguments: string);
var
  BtnActive: TToolbarButton97;
  okok: Boolean;
// d PT11
  LIJNettes, LIJBrutes, LGarantieNet: string;
  CALC: TBitBtn;
// f PT11
begin
  inherited;
  if LaTOB <> nil then
  begin
    TOB_Lignes := LaTOB;
  end;
  NETAPAYER := THNumEdit(GetControl('NETAPAYER'));

// d PT11
// calcul des IJSS
  if trim(Arguments) = 'IJSS' then
  begin
    Ecran.Caption := 'Calcul des IJSS';

    SetControlVisible('RUBDUBRUT', false);
    SetControlVisible('LBLRUB', false);

    SetControltext('NETAPAYER', FloatToStr(RendCumulSalSess('10')));
    SetControltext('LBLNET', 'Montant du ');
    SetControlEnabled('NETAPAYER', false);

    RechRubIJSS(RIJNettes, LIJNettes, RIJBrutes, LIJBrutes, RGarantieNet, LGarantieNet);
    if (RIJNettes = '') then
    begin
      PgiBox('Attention, La rubrique d''IJSS nettes n''a pas été paramétrée.#13#10' +
        'Cocher "Participe au calcul des IJSS"', 'Calcul des IJSS impossible.');
      SetControlEnabled('BVALIDER', FALSE);
      SetControlEnabled('CALC', FALSE);
    end;
    if (RIJBrutes = '') then
    begin
      PgiBox('Attention, La rubrique d''IJSS brutes n''a pas été paramétrée.#13#10' +
        'Cocher "Participe au calcul des IJSS"', 'Calcul des IJSS impossible.');
      SetControlEnabled('BVALIDER', FALSE);
      SetControlEnabled('CALC', FALSE);
    end;
    if (RGarantieNet = '') then
    begin
      PgiBox('Attention, La rubrique de garantie du net n''a pas été paramétrée.#13#10' +
        'Cocher "Participe au calcul des IJSS"', 'Calcul des IJSS impossible.');
      SetControlEnabled('BVALIDER', FALSE);
      SetControlEnabled('CALC', FALSE);
    end;
    if (RIJNettes = '') or (RIJBrutes = '') or (RGarantieNet = '') then
      Ecran.Close; //  PH Si IJSS non paramétré on sort

    SetControlVisible('CALC', true);

    SetControlVisible('RIJNETTES', true);
    SetControlVisible('LIJNETTES', true);
    SetControlVisible('MTIJNETTES', true);

    SetControlVisible('RIJBRUTES', true);
    SetControlVisible('LIJBRUTES', true);
    SetControlVisible('MTIJBRUTES', true);

    SetControlVisible('RGARANTIENET', true);
    SetControlVisible('LGARANTIENET', true);
    SetControlVisible('MTGARANTIENET', true);

    SetControltext('RIJNETTES', RIJNettes);
    SetControltext('LIJNETTES', LIJNettes);
    SetControltext('RIJBRUTES', RIJBrutes);
    SetControltext('LIJBRUTES', LIJBrutes);
    SetControltext('RGARANTIENET', RGarantieNet);
    SetControltext('LGARANTIENET', LGarantieNet);

    CALC := TBitBtn(GetControl('CALC'));
    if CALC <> nil then CALC.OnClick := CalcRubIJSS;

    MtIJNettes := THNumEdit(GetControl('MTIJNETTES'));
    MtIJBrutes := THNumEdit(GetControl('MTIJBRUTES'));
    MtGarantieNet := THNumEdit(GetControl('MTGARANTIENET'));

    BtnActive := TToolbarButton97(GetControl('BValider'));
    if BtnActive <> nil then BtnActive.OnClick := BValIJClick;
    Rubnet := False; // PT13-1
  end
  else
  begin
   // calcul de la paie à l'envers
   // DEB PT13-1
    if trim(Arguments) = 'RUB' then
    begin
      SetControlText('LBLNET', 'Saisissez le montant net de la rubrique s''ajoutant au ');
      SetControlProperty('NETAPAYER', 'LEFT', 134);
      SetControlProperty('NETAPAYER', 'TOP', 160);
      RubNet := TRUE;
    end
    else  Rubnet := False;
    // FINPT13-1
// f PT11

    RUBDUBRUT := THEdit(GetControl('RUBDUBRUT'));
// PT11  NETAPAYER := THNumEdit(GetControl('NETAPAYER'));
    if RUBDUBRUT <> nil then RUBDUBRUT.OnChange := RubChange;
    if (RUBDUBRUT <> nil) and (NETAPAYER <> nil) and (TOB_Lignes <> nil) then
    begin
     // PT1   : 16/01/2003 PH V591 Blocage calcul si pas de rem saisissable
      okok := RendRubPEnvers(TOB_Lignes); // Attention Tlig correspond à une TOB reelle du calcul du bulletin
      if Okok and (TLig <> nil) then RUBDUBRUT.Text := Copy(TLig.GetValue('PHB_RUBRIQUE'), 1, 4)
      else
      begin
        PgiBox('Attention, il n''a pas de rubrique saisissable dans le bulletin#13#10' +
          'insérez une rubrique saisissable dans votre bulletin', 'Calcul de la paie à l''envers');
        SetControlEnabled('BVALIDER', FALSE);
      end;
     // FIN PT1
    end;
    BtnActive := TToolbarButton97(GetControl('BValider'));
    if BtnActive <> nil then BtnActive.OnClick := BENVERSClick;
  end; // PT11

  SetControlText('LBLNET', GetControlText('LBLNET') + ' ' + RechDOM('PGMETHODEENVERS', VH_PAIE.PGENVERSNET, FALSE)); // PT8
end;

procedure TOF_PG_PaieEnvers.RubChange(Sender: TObject);
var
  T: TOB;
begin
  if RUBDUBRUT.Text = '' then exit;
  if (length(RUBDUBRUT.Text) < 4) or (length(RUBDUBRUT.Text) > 5) then exit;
  SetControlEnabled('BVALIDER', FALSE); // PT6 Bouton Validation/lancement désactivé
  T := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [RUBDUBRUT.Text], FALSE);
  if T = nil then
  begin
    PgiBox('Attention, La rubrique n''existe pas', 'Le Calcul de la paie à l''envers impossible');
    RUBDUBRUT.SetFocus;
    exit;
  end;
  if T.GetValue('PRM_CODECALCUL') = '01' then
  begin
    if (T.GetValue('PRM_TYPEMONTANT') <> '00') and (T.GetValue('PRM_TYPEMONTANT') <> '01') then
    begin
      PgiBox('Attention, Le montant de la rubrique n''est pas un élément permanent ou un élément variable', 'Le Calcul de la paie à l''envers impossible');
      RUBDUBRUT.SetFocus;
      exit;
    end;
  end;
  if T <> nil then TLig := TOB_Lignes.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', RUBDUBRUT.Text], FALSE);
  if (TLig = nil) or (T = nil) then
  begin
    PgiBox('Attention, La rubrique n''est pas présente dans le bulletin', 'Calcul de la paie à l''envers');
    RUBDUBRUT.SetFocus;
    Exit; // PT6
  end;
  SetControlEnabled('BVALIDER', TRUE); // PT6 si pas d'erreur alors calcul possible
end;

{ Fonction qui recherche dans la liste des lignes de la paie la premiere
remuneration qui est de type montant pour faire le calcul de la paie à l'envers
}

function TOF_PG_PaieEnvers.RendRubPEnvers(TOB_Lignes: TOB): Boolean;
var
  TRem: TOB;
  Rub: string;
begin
  Result := FALSE;
  TLig := TOB_Lignes.FindFirst([''], [''], FALSE);
  while TLig <> nil do
  begin
    Rub := Copy(TLig.GetValue('PHB_RUBRIQUE'), 1, 4);
    TRem := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [Rub], FALSE);
    if (TLig <> nil) and (TRem <> nil) then
    begin
      if TRem.GetValue('PRM_CODECALCUL') = '01' then
      begin
        if (TRem.GetValue('PRM_TYPEMONTANT') = '00') or (TRem.GetValue('PRM_TYPEMONTANT') = '01') then
        begin
          Result := True; //PT1  16/01/2003 PH V591 Blocage calcul si pas de rem saisissable
          break;
        end;
      end;
    end;
    TLig := TOB_Lignes.FindNext([''], [''], FALSE);
  end;
end;

{ Fonction activation du calcul de la paie à l'envers
}

procedure TOF_PG_PaieEnvers.BENVERSClick(Sender: TObject);
var
  St: string;
  OkOk: Boolean;
  PlusApprochant, Lenet: Double;  // PT13-1 variable utilisée lenet au lieu de NetApayer.text
begin
// DEB PT13-1
  if RubNet then
  begin
      // 1ere estimation du montant brut
    if VH_Paie.PGEnversNet = 'NET' then Lenet := RendCumulSalSess('10')
    else if VH_Paie.PGEnversNet = 'BRU' then lenet := RendCumulSalSess('01') // PT4
    else if VH_Paie.PGEnversNet = 'BRH' then lenet := RendCumulSalSess('05') // PT5
    else lenet := RendCumulSalSess('07') + RendCumulSalSess('08') + RendCumulSalSess('10');
    Lenet := Lenet + Valeur(NetAPayer.Text);
  end
  else Lenet := Valeur(NetAPayer.Text);
// FIN PT13-1
  OkOk := CalculPEnvers(PlusApprochant, TOB_Lignes, TLig, Lenet, TLIg.getvalue('PHB_RUBRIQUE'), FALSE); // PT13-1
  TFVierge(Ecran).Retour := 'N';
  if OkOk then TFVierge(Ecran).Retour := 'O'
  else
  begin // PT2   : 26/09/2003 PH V_421 FQ 10795 Si calcul n'a pas abaouti alors message indicatif
  // PT7 On relance le calcul en gérant l'arrondi
    if (not OkOk) and (VH_PAIE.PGArrondiEnvers > 0) then
      OkOk := CalculPEnvers(PlusApprochant, TOB_Lignes, TLig, Lenet, TLIg.getvalue('PHB_RUBRIQUE'), FALSE, TRUE); // PT13-1
    if not OkOk then
    begin
      st := 'Le calcul n''a pas abouti, mais la valeur la plus proche #13#10' +
        'correspondant à vos critères est de ' + DoubleToCell(PlusApprochant, 2) +
        '#13#10soit un écart de ' + DoubleToCell(Lenet - PlusApprochant, 2);      // PT13-1
      PgiBox(St, Ecran.Caption);
    end;
  end;

end;
{$ENDIF}

{ Fonction initialisation des bornes de calcul pour la paie à l'envers
}

procedure PEnversRechercheValeur(var ajust, netinf, netsup, brutinf, brutsup, salaire_estime, net_a_payer, net: double; var Ziteration: Integer);
var
  a, b: double;
begin
  a := 0;
  b := 0;
  { si netinf == net_a_payer ou netsup == net_a_payer, alors
    on peut directement affecter la borne du brut correspondant}
  if netinf = net_a_payer then salaire_estime := brutinf
  else
  begin // A
    if netsup = net_a_payer then salaire_estime := brutsup
    else
    begin // B
      if (netinf = 0) or (netsup = 0) then
      begin
        Ziteration := 1;
        if netinf = 0 then salaire_estime := brutsup + ajust;
        if netsup = 0 then salaire_estime := brutinf + ajust;
      end
      else // C
      begin
        // Logique d'interpolation
        if (Ziteration >= 1) and (Ziteration <= 2) then
        begin
          if brutsup <> brutinf then
          begin
            a := (netsup - netinf) / (brutsup - brutinf);
            b := (netsup + netinf - ((netsup - netinf) / (brutsup - brutinf)) * (brutsup + brutinf)) / 2;
            Ziteration := Ziteration + 1;
          end;
          if a <> 0 then salaire_estime := (net_a_payer - b) / a;
        end
          // recherche dichotomique
        else salaire_estime := (brutinf + brutsup) / 2;
      end // Fin C
    end // fin B
  end; // fin A

end;
{ Fonction initialisation des bornes de calcul pour la paie à l'envers
}

procedure BEnversInit(var netinf, netsup, net, net_a_payer, brutinf, brutsup, montant: double);
begin
  if net < net_a_payer then
  begin
    netinf := net;
    brutinf := montant;
  end
  else
  begin
    netsup := net;
    brutsup := montant;
  end;
end;

function DonneSens(Rub, Cumul: string): string;
var
  T_Sens, TRechCumRub: TOB;
  I: Integer;
  Cum: string;
begin
  Result := '';
  T_Sens := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [Rub], FALSE);
  if T_Sens <> nil then
  begin
    for I := 0 to T_Sens.Detail.Count - 1 do // boucle sur la recherche des cumuls alimentés par la rubrique
    begin
      TRechCumRub := T_Sens.Detail[I];
      Cum := TRechCumRub.GetValue('PCR_CUMULPAIE');
      if Cumul = Cum then Result := TRechCumRub.GetValue('PCR_SENS');
    end;
  end;
end;

function RendSensRub(Rub: string): string;
begin
  Result := '';
  Result := DonneSens(Rub, '01');
  if Result = '' then
  begin
    Result := DonneSens(Rub, '02');
    if Result = '' then
    begin
      Result := DonneSens(Rub, '09');
      if Result = '' then DonneSens(Rub, '10');
    end;
  end;
  // if Result = '' then Sens := '+';
end;
{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 18/05/2005
Modifié le ... :   /  /
Description .. : Fonction publique calcul de la paie à l'envers
Suite ........ :
Suite ........ : PT7 rajout paramètre pour gestion de l'arrondi
Mots clefs ... :
*****************************************************************}
// PT9

function CalculPEnvers(var PlusApprochant: Double; TOB_Lignes, TLig: TOB; NetTrouve: Double; Rub: string; Silencieux: Boolean; GestArrondi: Boolean = FALSE; CalcMaint: Boolean = FALSE): Boolean;
var
  montant: double;
  okok: Boolean;
  B1, B2: Double;
  Salaire_estime, netinf, netsup, net, net_a_payer, brutinf, brutsup, ajust: Double; // variables
  Ziteration, nb_iteration, MAX_iteration: integer; // compteur sur les iterations
  Sens: string;
begin
  PlusApprochant := 0;
  Ziteration := 0;
  nb_iteration := 0;
  MAX_iteration := VH_Paie.PGIterationEnvers;
  Salaire_estime := 0;
  netinf := 0;
  netsup := 0;
  net := 0;
  net_a_payer := 0;
  brutinf := 0;
  brutsup := 0;
  ajust := 0;
  net_a_payer := NetTrouve;
  Sens := RendSensRub(Rub);
  // 1ere estimation du montant brut
  if VH_Paie.PGEnversNet = 'NET' then net := RendCumulSalSess('10')
  else if VH_Paie.PGEnversNet = 'BRU' then net := RendCumulSalSess('01') // PT4
  else if VH_Paie.PGEnversNet = 'BRH' then net := RendCumulSalSess('05') // PT5
  else net := RendCumulSalSess('07') + RendCumulSalSess('08') + RendCumulSalSess('10');

  salaire_estime := TLig.GetValue('PHB_MTREM') + ((net_a_payer - net) * 1.2);
  TLig.PutValue('PHB_MTREM', salaire_estime);
  // Rajouter ici si on traite un autre code calcul que 01 par ex calculer la base
//PT9  CalculBulletin(TOB_Lignes);
  CalculBulletin(TOB_Lignes, CalcMaint, nil, True); // PT10
  if VH_Paie.PGEnversNet = 'NET' then net := RendCumulSalSess('10')
  else if VH_Paie.PGEnversNet = 'BRU' then net := RendCumulSalSess('01') // PT4
  else if VH_Paie.PGEnversNet = 'BRH' then net := RendCumulSalSess('05') // PT5
  else net := RendCumulSalSess('07') + RendCumulSalSess('08') + RendCumulSalSess('10');
  // Initialisation des bornes de calcul pour des recherches par itérations
  montant := TLig.GetValue('PHB_MTREM');
  BEnversInit(netinf, netsup, net, net_a_payer, brutinf, brutsup, montant);
  if not Silencieux then
    InitMoveProgressForm(nil, 'Calcul de la PAIE à l''envers', 'Calcul en cours, Veuillez patienter SVP ...', MAX_iteration + 1, TRUE, TRUE);
  okok := TRUE;
  // boucle de recherche par iteration limitee à un maximum pour eviter une boucle sans fin
  while ((net_a_payer <> net) and (nb_iteration <> MAX_iteration)) do
  begin
// PT7 gestion de l'arrondi
    if (VH_Paie.PGArrondiEnvers > 0) and GestArrondi then
    begin
      if (ABS(net_a_payer - net) <= VH_Paie.PGARRONDIENVERS) then break;
    end
    else if net_a_payer = net then break;
// FIN PT7
    if nb_iteration = MAX_iteration then break;
    nb_iteration := nb_iteration + 1;
    if not Silencieux then
      okok := MoveCurProgressForm(IntToStr(nb_iteration));
    if not OkOk then break;
    if (netinf = 0) or (netsup = 0) then
    begin
      if netinf = 0 then ajust := ((net_a_payer - netsup) * 1.2);
      if netsup = 0 then ajust := ((net_a_payer - netinf) * 1.2);
      if Sens = '-' then ajust := ajust * -1;
    end;
    PEnversRechercheValeur(ajust, netinf, netsup, brutinf, brutsup, salaire_estime, net_a_payer, net, Ziteration);
    // Arrondi
    salaire_estime := ARRONDI(salaire_estime, 2);
    Rub := TLIg.getvalue('PHB_RUBRIQUE');
    TLig.PutValue('PHB_MTREM', salaire_estime);
//PT9  CalculBulletin(TOB_Lignes);
    CalculBulletin(TOB_Lignes, CalcMaint, nil, True); // PT10
    if VH_Paie.PGEnversNet = 'NET' then net := RendCumulSalSess('10')
    else if VH_Paie.PGEnversNet = 'BRU' then net := RendCumulSalSess('01') // PT4
    else if VH_Paie.PGEnversNet = 'BRH' then net := RendCumulSalSess('05') // PT5
    else net := RendCumulSalSess('07') + RendCumulSalSess('08') + RendCumulSalSess('10');
    net := ARRONDI(net, 2);
    // PT2   : 26/09/2003 PH V_421 FQ 10795 Si calcul n'a pas abaouti alors message indicatif
    if PlusApprochant = 0 then PlusApprochant := net
    else
    begin
      B1 := Abs(net_a_payer - PlusApprochant);
      B2 := Abs(net - net_a_payer);
      if B1 > B2 then PlusApprochant := net;
    end;
    // FIN PT2
    // Initialisation des bornes de calcul pour des recherches par itérations
    montant := TLig.GetValue('PHB_MTREM');
    BEnversInit(netinf, netsup, net, net_a_payer, brutinf, brutsup, montant);
  end;
  if not Silencieux then FiniMoveProgressForm();
  if (net_a_payer = net) then Result := True
  else Result := FALSE;
end;
// d PT11
{***********A.G.L.***********************************************
Auteur  ...... : MF
Créé le ...... : 03/05/2007
Modifié le ... :   /  /
Description .. : Fonction de récupération des n° de rubriques
Suite ........ : participant au calcul des IJSS
Suite ........ : Utilisation du champ PRM_DADS
Suite ........ : pos 1 = booléen (participe True ou False)
Suite ........ : pos 2 à 4 = Type de la rubrique (IJ nettes, Ij brutes ou Garantie)
Mots clefs ... : IJSS
*****************************************************************}
{$IFNDEF EAGLSERVER}

procedure TOF_PG_PaieEnvers.RechRubIJSS(var RIJNettes, LIJNettes, RIJBrutes, LIJBrutes, RGarantieNet, LGarantieNet: string);
var
  Q: TQuery;
  st: string;
begin
  RIJNettes := '';
  RIJBrutes := '';
  RGarantieNet := '';
  LIJNettes := '';
  LIJBrutes := '';
  LGarantieNet := '';

  Q := OpenSQL('select PRM_RUBRIQUE, PRM_LIBELLE, PRM_DADS from remuneration ' +
    'where PRM_DADS like "X%"', TRUE);
  while not Q.EOF do
  begin
    st := Q.FindField('PRM_DADS').AsString;
    if copy(st, 2, 3) = '001' then
    // IJSS nettes
    begin
      RIJNettes := Q.FindField('PRM_RUBRIQUE').AsString;
      LIJNettes := Q.FindField('PRM_LIBELLE').AsString;
    end
    else
      if copy(st, 2, 3) = '002' then
    // IJSS brutes
      begin
        RIJBrutes := Q.FindField('PRM_RUBRIQUE').AsString;
        LIJBrutes := Q.FindField('PRM_LIBELLE').AsString;
      end
      else
        if copy(st, 2, 3) = '003' then
        begin
    // Garantie du net
          RGarantieNet := Q.FindField('PRM_RUBRIQUE').AsString;
          LGarantieNet := Q.FindField('PRM_LIBELLE').AsString;
        end;
    if (RIJNettes = '') or (RIJBrutes = '') or (RGarantieNet = '') then
      Q.Next
    else
      break;
  end;
  Ferme(Q);
end; // fin  RechRubIJSS

{***********A.G.L.***********************************************
Auteur  ...... : MF
Créé le ...... : 03/05/2007
Modifié le ... :   /  /
Description .. : Fonction de calcul du montant des IJSS brutes et
Suite ........ : du montant de la garantie du net à payer (paie à
Suite ........ : l'envers)
Mots clefs ... : IJSS
*****************************************************************}

procedure TOF_PG_PaieEnvers.CalcRubIJSS(Sender: TObject);
var
  RetEnvers: boolean;
  Sauv_PGEnversNet, st: string;
  PlusApprochant: double;
begin
  SourisSablier; // PT12
  if (Valeur(MtIJNettes.Text) <> 0) then
  begin
    Tlig := TOB_Lignes.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', RIJNettes], FALSE);
    if TLig = nil then
    begin
      // on insere la rubrique dans le bulletin
      ChargeProfilSPR(TOB_Salarie, TOB_Lignes, RIJNettes, 'AAA');
      Tlig := TOB_Lignes.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', RIJNettes], FALSE);
    end;
       // valeur de la ligne pour le calcul de la paie à l'envers
    if TLig <> nil then
      TLig.putValue('PHB_MTREM', Valeur(MtIJNettes.Text));

      // calcul des IJSS brutes
    if (ValEltNat('0098', TLig.GetValue('PHB_DATEFIN'), nil) = 0) then
    begin
      PgiBox('Le taux réduit de CSG/CRDS applicable aux revenus de remplacemnt #13#10' +
        'n''a pas été renseigné : Elément national 0098', 'Calcul des IJSS brutes.');
    end;

    MtIJBrutes.text := FloatToStr(ARRONDI(Valeur(MtIJNettes.Text) / (1 - (ValEltNat('0098', TLig.GetValue('PHB_DATEFIN'), nil) / 100)), 2));

    Tlig := TOB_Lignes.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', RIJBrutes], FALSE);
    if TLig = nil then
    begin
      // on insere la rubrique dans le bulletin
      ChargeProfilSPR(TOB_Salarie, TOB_Lignes, RIJBrutes, 'AAA');
      Tlig := TOB_Lignes.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', RIJBrutes], FALSE);
    end;
       // valeur de la ligne pour le calcul de la paie à l'envers
    if TLig <> nil then
      TLig.putValue('PHB_MTREM', Valeur(MtIJBrutes.Text));

      // on insère la rubrique de garantie dans le bulletin
    ChargeProfilSPR(TOB_Salarie, TOB_Lignes, RGarantieNet, 'AAA');
    Tlig := TOB_Lignes.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', RGarantieNet], FALSE);
       // valeur (0) de la ligne pour le calcul de la paie à l'envers
    if TLig <> nil then
      TLig.putValue('PHB_MTREM', 0);


    if (Valeur(NETAPAYER.Text) <> 0) and (RGarantieNet <> '') then
    begin
      if TLig <> nil then
      begin
        PlusApprochant := 0;

        Sauv_PGEnversNet := VH_Paie.PGEnversNet;
        VH_Paie.PGEnversNet := 'NET';

        RetEnvers := CalculPEnvers(PlusApprochant, TOB_Lignes, TLig, Valeur(NETAPAYER.Text), RGarantieNet, TRUE, FALSE, FALSE);
           //  On relance le calcul avec arrondi si echec
        if (not RetEnvers) and (VH_PAIE.PGArrondiEnvers > 0) then
          RetEnvers := CalculPEnvers(PlusApprochant, TOB_Lignes, TLig, Valeur(NETAPAYER.Text), RGarantieNet, TRUE, TRUE, FALSE);
        if not RetEnvers then
        begin
          st := 'Le calcul n''a pas abouti, mais la valeur la plus proche #13#10' +
            'correspondant à vos critères est de ' + DoubleToCell(PlusApprochant, 2) +
            '#13#10soit un écart de ' + DoubleToCell(Valeur(NetAPayer.Text) - PlusApprochant, 2);
          PgiBox(St, Ecran.Caption);
        end;

        VH_Paie.PGEnversNet := Sauv_PGEnversNet;

        Tlig := TOB_Lignes.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', RGarantieNet], FALSE);
        if TLig <> nil then
        begin
          MtGarantieNet.text := FloatToStr(TLig.getValue('PHB_MTREM'));
        end;
      end;
    end;

  end
  else
  begin
    PgiBox('Le montant des IJSS nettes n''a pas été saisi.');
    MtIJNettes.SetFocus;
    exit;
  end;

  SourisNormale; // PT12
  SetControlEnabled('BVALIDER', TRUE); //  si pas d'erreur alors calcul possible
end;
{***********A.G.L.***********************************************
Auteur  ...... : MF
Créé le ...... : 03/05/2007
Modifié le ... :   /  /
Description .. : Traitement du bouton VALIDER
Mots clefs ... : IJSS
*****************************************************************}

procedure TOF_PG_PaieEnvers.BValIJClick(Sender: TObject);
begin
  TFVierge(Ecran).Retour := 'O'
end;
{$ENDIF}
// f PT11

initialization
{$IFNDEF EAGLSERVER}
  registerclasses([TOF_PG_PaieEnvers]);
{$ENDIF}
end.

