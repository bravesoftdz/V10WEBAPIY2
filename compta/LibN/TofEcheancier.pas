{Unité : TofEcheancier , Fiche : EFECHEANCIER (QRS1)
--------------------------------------------------------------------------------------
  Version     |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
06.51.001.xxx  13/10/05   JP   FQ 16861 : branchement de l'état depuis le détail des mouvements
07.00.001.003  28/02/06   JP   FQ 17263 : remise à plat de la gestion des généraux et des Tiers
                               Gestion plus calssique (!) de la constitution de la requête dans le OnLoad
--------------------------------------------------------------------------------------}
unit TofEcheancier;

interface

uses Classes, StdCtrls,dialogs,
{$IFDEF EAGLCLIENT}
     eQRS1,MaineAGL,
{$ELSE}
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     QRS1,
     FE_Main,
{$ENDIF}
     comctrls,UTof, HCtrls,Ent1,  TofMeth, HTB97, spin,
     filtre,
     TofTLMvtRupt,
     TofTLMvtTri,
     SaisUtil,
     SysUtils,
     ParamSoc,		//GetParamSocSecur YMO
     Windows ;

procedure CC_LanceFicheEcheancier; // Modif SBO Fusion VDEV

type
    TOF_ECHEANCIER = class(TOF_Meth)
  private
    Devise          : THValComboBox ;
    NatureAuxi      : THValComboBox ;
    NatureGene      : THMultiValComboBox;
    SensMouvement   : THValComboBox ;
    TypeEdition     : THValComboBox ;
    Paiement        : THValComboBox ;
    Reference       : THValComboBox ;
    Libelle         : THValComboBox ;
    Libelle2        : THValComboBox ;
    FFiltres        : THValComboBox ;
    Affichage       : THRadioGroup ;
    Impression      : THRadioGroup ;
    StDebit         : String ;
    StCredit        : String ;
    NatureAux       : String ;
    NatureGen       : String ;
    AuxInf, AuxSup  : THEdit ;
    ColInv, ColVis  : THEdit ;
    GenInf, GenSup  : THEdit ;
    procedure NatureAuxiOnChange (Sender : TObject) ;
    procedure NatureGeneOnChange (Sender : TObject) ;
    procedure SensMouvementOnChange (Sender : TObject) ;
    procedure DeviseOnChange (Sender : TObject);
    procedure Rupture1OnDblClick (Sender : TObject);
    procedure Rupture1OnKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure TriParOnDblClick (Sender : TObject);
    procedure AffichageOnClick (Sender : TObject);
    procedure ImpressionOnClick (Sender : TObject);
    procedure TriParOnChange (Sender : TObject);
    // Gestion des filtres
    procedure InitZones ;
    procedure NouvRechClick(Sender:TObject) ;
    procedure SupprFiltreClick(Sender:TObject) ;
    procedure AuxiElipsisClick         ( Sender : TObject );

  public
    procedure OnArgument(Arguments : string); override ;
    procedure OnUpdate; override ;
    procedure OnLoad; override;{JP 28/02/05 : FQ 17263}
    procedure OnNew ; override ;
    procedure OnAfterFormShow; override; {JP 29/06/06}
    procedure ChargementCritEdt; override; {JP 13/10/05 : FQ 16861}
    procedure ColVisKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState); {JP 28/02/05 : FQ 17263}
    procedure GereComptesExceptions; {FQ 17263 : touche "F5" pour les comptes d'exceptions}
  end ;

implementation

uses
  AGLInit, {JP 13/10/05 : FQ 16861}
  CritEdt,{JP 13/10/05 : FQ 16861}
  ULibWindows, {JP 27/02/06 : FQ 17263 : TraductionTHMultiValComboBox}
  HQry, { BVE 18.04.07 FQ 18007}
  HEnt1, LicUtil,HMsgBox, UTofMulParamGen;


procedure CC_LanceFicheEcheancier;
begin
  AGLLanceFiche('CP','EPECHEANCIER','','','');
end;

procedure TOF_ECHEANCIER.OnUpdate;
begin
  inherited;
//  TFQRS1(Ecran).WhereSQL:='|'+Order;
end;

PROCEDURE TOF_ECHEANCIER.OnArgument(Arguments : string);
VAR
  Rupture1,Rupture2 : THEdit;
  TriPar            : THEdit;
  AffichageDev      : THRadioGroup;
  DeviseAff         : RDevise;
BEGIN
  inherited ;
  { FQ 20039 BVE 25.04.07 }
  AffichageDev := THRadioGroup(GetControl('AFFICHAGE'));
  DeviseAff.Code := GetParamSoc('SO_DEVISEPRINC');
  GetInfosDevise(DeviseAff);
  if DeviseAff.Libelle <> '' then
     AffichageDev.Items.Strings[0] := DeviseAff.Libelle;
  { END FQ 20039 }

  with TFQRS1(Ecran) do
  begin
  {JP 19/08/05 : FQ 16479 : Je ne vois pas l'intérêt d'appeler la fonction car tous les
                 tags sont à 0 et cela change l'ordre des onglets
  PageOrder(Pages);}
  NatureEtat := Arguments;
{$IFDEF EAGLCLIENT}
  ChoixEtat:=TRUE ;
{$ENDIF}
  end;
StDebit :='-(E_COUVERTURE*((E_DEBIT+E_DEBITDEV)/(E_DEBIT+E_DEBITDEV+E_CREDIT+E_CREDITDEV))))';
StCredit:='-(E_COUVERTURE*((E_CREDIT+E_CREDITDEV)/(E_DEBIT+E_DEBITDEV+E_CREDIT+E_CREDITDEV))))';
NatureAuxi:=THValComboBox(GetControl('T_NATUREAUXI'));
NatureGene:=THMultiValComboBox(GetControl('G_NATUREGENE'));
SensMouvement:=THValComboBox(GetControl('SENSMOUVEMENT'));
TypeEdition:=THValComboBox(GetControl('TYPEEDITION'));
Devise:=THValComboBox(GetControl('E_DEVISE'));
Paiement:=THValComboBox(GetControl('PAIEMENT'));
Reference:=THValComboBox(GetControl('REFERENCE'));
Libelle:=THValComboBox(GetControl('LIBELLE'));
Libelle2:=THValComboBox(GetControl('LIBELLE2'));
Rupture1:=THEdit(GetControl('RUPTURE1'));
Rupture2:=THEdit(GetControl('RUPTURE2'));
TriPar:=THEdit(GetControl('TRIPAR'));
AuxInf:=THEdit(GetControl('E_AUXILIAIRE')) ;
AuxSup:=THEdit(GetControl('E_AUXILIAIRE_')) ;
GenInf:=THEdit(GetControl('E_GENERAL')) ;
GenSup:=THEdit(GetControl('E_GENERAL_')) ;
Affichage:=THRadioGroup(GetControl('AFFICHAGE'));
Impression:=THRadioGroup(GetControl('TIMPRESSION'));
FFiltres := THValComboBox(GetControl('FFILTRES'));
TFQRS1(Ecran).OnAfterFormShow := OnAfterFormShow ;
THValComboBox(GetControl('TYPEEDITION')).ItemIndex:=0;
SetControlText('E_DATEECHEANCE',DateToStr(now));
if now<VH^.Encours.Fin then
   SetControlText('E_DATEECHEANCE_',DateToStr(VH^.Encours.Fin))
   else SetControlText('E_DATEECHEANCE_',DateToStr(VH^.Suivant.Fin));

  if Devise <> nil then begin
    Devise.OnChange := DeviseOnChange;
    Devise.itemindex := 0;
    Devise.OnChange(Devise);
  end;
  if NatureAuxi <> nil then NatureAuxi.OnChange := NatureAuxiOnChange;
  if NatureGene <> nil then NatureGene.OnChange := NatureGeneOnChange;

if TypeEdition<>nil then TypeEdition.itemindex:=0;
if SensMouvement<>nil then
   BEGIN
   SensMouvement.OnChange:=SensMouvementOnChange;
   SensMouvement.itemindex:=2;
   SetControlText('SENS','(('+(GetControlText('DEBIT'))+StDebit+'-('+(GetControlText('CREDIT'))+StCredit+')');
   END ;
if Affichage<>nil then
  BEGIN
  Affichage.OnClick:=AffichageOnClick;
  Affichage.itemindex:=0;
  SetControlText('DEBIT','E_DEBIT');
  SetControlText('CREDIT','E_CREDIT');
  AffichageOnClick(nil);
  END ;
if Impression<>nil then
  BEGIN
  Impression.OnClick:=ImpressionOnClick;
  Impression.itemindex:=0;
  SetControlText('IMPRESSION','DECI');
  END ;
if Rupture1<>nil then
begin
  Rupture1.OnDblClick:=Rupture1OnDblClick;
  Rupture1.OnKeyDown:=Rupture1OnKeyDown;
end;
if Rupture2<>nil then
begin
  Rupture2.OnDblClick:=Rupture1OnDblClick;
  Rupture2.OnKeyDown:=Rupture1OnKeyDown;
end;

if TriPar<>nil   then
  BEGIN
  TriPar.OnChange:=TriParOnChange;
  TriPar.OnDblClick:=TriParOnDblClick;
  TriPar.OnKeyDown:=Rupture1OnKeyDown;
  SetControlEnabled('SAUTPAGE',False);
  SetControlEnabled('SURRUPTURE',False);
  END ;
if Reference<>nil then THValComboBox(GetControl('REFERENCE')).ItemIndex:=0 ;
if Libelle<>nil   then THValComboBox(GetControl('LIBELLE')).ItemIndex:=2 ;
if Libelle2<>nil  then THValComboBox(GetControl('LIBELLE2')).ItemIndex:=3 ;
if Paiement<>nil  then THValComboBox(GetControl('PAIEMENT')).ItemIndex:=4 ;

  {FQ 17263 : Control invisible utilisé pour le Lookup}
  ColInv := THEdit(GetControl('CPTEXCEPT'));
  {FQ 17263 : Control contenant les comptes généraux à exclure du traitement, séparés par un ";"}
  ColVis := THEdit(GetControl('COMPTESEXCLUS'));
  {JP 28/02/06 : FQ 17263 : Branchement du F5 sur les comptes à exclure}
  if Assigned(ColVis) then
    ColVis.OnKeyDown := ColVisKeyDown;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    THEdit(GetControl('E_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;
    THEdit(GetControl('E_AUXILIAIRE_', true)).OnElipsisClick:=AuxiElipsisClick;
  end;

END;

PROCEDURE TOF_ECHEANCIER.Rupture1OnDblClick (Sender : TObject);
VAR    St,St1,St2,St3,Arg:String;
       i : integer;
       OkAssoc : integer;
BEGIN
OkAssoc:=0;
Arg:=GetControlText('RUPTURE1')+'|'+GetControlText('RUPTURE2');
St:=CPLanceFiche_MvtRuptureTL('','',Arg);
St1 :=ReadTokenPipe(St,'|');
SetControlText('RUPTURE1',St1);
SetControlText('RUPTURE2',St);

if St1 <> '' then
  for i:=0 to 3 do
    BEGIN
    St2:=ReadTokenSt(St1);
    if (St2<>'#') and (St2<>'-') then BEGIN St3:=St3+'E0'+IntToStr(i)+';'; OkAssoc:=OkAssoc+1 ; END ;
    END
else St3 := '' ;
SetControlText('TRIPAR',St3);

if OkAssoc=0 then
  BEGIN
  SetControlEnabled('CPTASSOCIES',False);
  SetControlChecked('CPTASSOCIES',False);
  END else
  BEGIN
  SetControlEnabled('CPTASSOCIES',True);
  SetControlChecked('CPTASSOCIES',True);
  END ;

END;


procedure TOF_ECHEANCIER.Rupture1OnKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 case Key of
    VK_DELETE, VK_BACK    : begin
                              SetControlText('RUPTURE1','');
                              SetControlText('RUPTURE2','');
                              SetControlText('TRIPAR','');
                            end;
 end;
end;

PROCEDURE TOF_ECHEANCIER.TriParOnChange (Sender : TObject);
BEGIN
if GetControlText('TRIPAR')='' then
  begin
  SetControlChecked('SAUTPAGE',False);
  SetControlChecked('SURRUPTURE',False);
  SetControlEnabled('SAUTPAGE',False);
  SetControlEnabled('SURRUPTURE',False);
  end else
  begin
  SetControlEnabled('SAUTPAGE',True);
  SetControlEnabled('SURRUPTURE',True);
  end;
END;

PROCEDURE TOF_ECHEANCIER.TriParOnDblClick (Sender : TObject);
VAR    St,Arg:String;
BEGIN
Arg:=GetControlText('TRIPAR');
St:=CPLanceFiche_MvtTriTL('','',Arg);
SetControlText('TRIPAR',St);
END;

PROCEDURE TOF_ECHEANCIER.ImpressionOnClick (Sender : TObject);
BEGIN
case Impression.ItemIndex of
  0 : SetControlText('IMPRESSION','DECI');
  1 : SetControlText('IMPRESSION','KILO');
  2 : SetControlText('IMPRESSION','MEGA');
  3 : SetControlText('IMPRESSION','ENTI');
  END ;
END;

procedure TOF_ECHEANCIER.AffichageOnClick (Sender : TObject);
begin
  if (Affichage.ItemIndex=1) and ((GetControlText('E_DEVISE') = V_PGI.DevisePivot)) then
  begin
    PgiBox('Vous devez d''abord sélectionner une devise particulière','Echéancier');
    THRadiogroup(GetControl('AFFICHAGE')).ItemIndex:=0;
  end;

  case Affichage.ItemIndex of
    0 : begin
          SetControlText('DEBIT' , 'E_DEBIT');
          SetControlText('CREDIT', 'E_CREDIT');
        end;
    1 : begin
          SetControlText('DEBIT' , 'E_DEBITDEV');
          SetControlText('CREDIT', 'E_CREDITDEV');
        end;
  end;
  SensMouvementOnChange(nil);
END;

procedure TOF_ECHEANCIER.DeviseOnChange (Sender : TObject);
begin
  if (Devise.value = V_PGI.DevisePivot) and
     (THRadioGroup(GetControl('AFFICHAGE')).ItemIndex = 1) then begin
    PgiBox('Cette devise n''est pas compatible avec l''affichage sélectionné','Echéancier');
    THRadiogroup(GetControl('AFFICHAGE')).ItemIndex := 0;
  end;

  if (THValComboBox(GetControl('E_DEVISE')).ItemIndex = 0) and
     (THRadioGroup(GetControl('AFFICHAGE')).ItemIndex = 1) then
    THRadiogroup(GetControl('AFFICHAGE')).ItemIndex := 0;
end;

PROCEDURE TOF_ECHEANCIER.SensMouvementOnChange (Sender : TObject);
BEGIN
case SensMouvement.ItemIndex of
  0 : SetControlText('SENS','(('+GetControlText('CREDIT')+StCredit+'*-1)');
  1 : SetControlText('SENS','(('+GetControlText('DEBIT')+StDebit+')');
  2 : SetControlText('SENS','(('+GetControlText('DEBIT')+StDebit+'-('+GetControlText('CREDIT')+StCredit+')');
  END ;
END;

procedure TOF_ECHEANCIER.InitZones;
begin
  NatureAux := '';
  NatureGen := '';

  TypeEdition.itemindex:=0;
  SensMouvement.itemindex:=2;
  SetControlText('SENS','(('+(GetControlText('DEBIT'))+StDebit+'-('+(GetControlText('CREDIT'))+StCredit+')');

  Affichage.itemindex:=0;
  SetControlText('DEBIT','E_DEBIT');
  SetControlText('CREDIT','E_CREDIT');
  // YMO 09/01/2006 FQ 17240 La fonction Videfiltre réinitialise les zones
  SetControlText('E_QUALIFPIECE', 'N;'); {FQ 17263 : remplacement de toutes les zones par un MultiValCombo}
end;

procedure TOF_ECHEANCIER.NouvRechClick(Sender: TObject);
begin
  videFiltre( FFiltres, Pages ) ;
  TFQRS1(Ecran).ListeFiltre.new ;

  // === INIT ZONES
  InitZones ;
end;

procedure TOF_ECHEANCIER.SupprFiltreClick(Sender: TObject);
begin
  if not TFQRS1(Ecran).ListeFiltre.Delete then Exit ;
  videFiltre( FFiltres, Pages ) ;

  InitZones ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 12/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_ECHEANCIER.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;


procedure TOF_ECHEANCIER.OnNew;
begin
  inherited;

  // Initialisation des champs si pas de filtres
  if FFiltres.Text = '' then
    {JP 15/10/07 : FQ 16149 : Le InitZones suffit largement en ouverture de fiche}
      InitZones ;
//NouvRechClick( nil ) ;

end;

procedure TOF_ECHEANCIER.OnAfterFormShow;
begin
  TFQRS1(Ecran).ListeFiltre.OnItemNouveau   := NouvRechClick ;
  TFQRS1(Ecran).ListeFiltre.OnItemSupprimer := SupprFiltreClick ;
end;

{JP 13/10/05 : FQ 16861 : Branchement de l'état en eAgl depuis le détail des mouvements
{---------------------------------------------------------------------------------------}
procedure TOF_ECHEANCIER.ChargementCritEdt;
{---------------------------------------------------------------------------------------}
begin
  inherited;

  if (TheData <> nil) and (TheData is ClassCritEdt) then begin
    SetControlText('E_AUXILIAIRE' , ClassCritEdt(TheData).CritEdt.Cpt1);
    SetControlText('E_AUXILIAIRE_', ClassCritEdt(TheData).CritEdt.Cpt2);

    SetControlText('E_GENERAL' , ClassCritEdt(TheData).CritEdt.Ech.Cpt3);
    SetControlText('E_GENERAL_', ClassCritEdt(TheData).CritEdt.Ech.Cpt4);

    SetControlText('E_JOURNAL' , ClassCritEdt(TheData).CritEdt.Ech.Jal1);
    SetControlText('E_JOURNAL_', ClassCritEdt(TheData).CritEdt.Ech.Jal2);

    SetControlText('E_DATEECHEANCE' , DateToStr(ClassCritEdt(TheData).CritEdt.Date1));
    SetControlText('E_DATEECHEANCE_', DateToStr(ClassCritEdt(TheData).CritEdt.Date2));

    SetControlText('E_DEVISE', ClassCritEdt(TheData).CritEdt.DeviseSelect);
    SetControlText('E_ETABLISSEMENT', ClassCritEdt(TheData).CritEdt.Etab);

    SetControlText('E_QUALIFPIECE', ClassCritEdt(TheData).CritEdt.Qualifpiece);
    TheData := nil;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHEANCIER.ColVisKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  case Key of
    VK_F5 : GereComptesExceptions;{FQ 17263}
  end;
end;

{Simulation du F5 sur le THEdit contenant les comptes généraux à exclure (ColVis) en utilisant
 un THedit invisible (ColInv) / FQ 17263
{---------------------------------------------------------------------------------------}
procedure TOF_ECHEANCIER.GereComptesExceptions;
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  ColInv.Text := '';
  ColInv.ElipsisClick(ColInv);
  if ColInv.Text <> '' then begin
    s := ColVis.Text;
    if (s <> '') and (s[Length(s)] <> ';') then s := s + ';';
    s := s + ColInv.Text + ';';
    ColVis.Text := s;
    ColVis.SetFocus;
  end;
end;

{FQ 17263 : Refonte de la gestion des comptes généraux et auxiliaires
{---------------------------------------------------------------------------------------}
procedure TOF_ECHEANCIER.NatureAuxiOnChange (Sender : TObject) ;
{---------------------------------------------------------------------------------------}
var
  SQL : string;
begin
  NatureAux := GetControlText('T_NATUREAUXI') ;
  { FQ 18571 BVE 19.04.07 }
  if not (cLoadFiltre) then
  begin
     SetControlText('E_AUXILIAIRE', '');
     SetControlText('E_AUXILIAIRE_', '');
  end;

  {Consitution de la clause plus}
  if NatureAux <> '' then
     SQL := 'T_NATUREAUXI = "' + NatureAux + '"'
  else
     SQL := '';      
  { END FQ 18751 }
  AuxInf.Plus := SQL;
  AuxSup.Plus := SQL;
end;

{FQ 17263 : Refonte de la gestion des comptes généraux et auxiliaires
{---------------------------------------------------------------------------------------}
procedure TOF_ECHEANCIER.NatureGeneOnChange (Sender : TObject) ;
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  Lib : string;
begin
  NatureGen := GetControlText('G_NATUREGENE') ;

  {On commence par vider les zones}
  SetControlText('COMPTESEXCLUS', '');
  SetControlText('CPTEXCEPT', '');
  { FQ 18571 BVE 19.04.07 }
  if not(cLoadFiltre) then
  begin
     SetControlText('E_GENERAL','');
     SetControlText('E_GENERAL_','');
  end;
  { END FQ 18571 }

  {Consitution de la clause plus}
  TraductionTHMultiValComboBox(NatureGene, SQL, Lib, 'G_NATUREGENE');

  {Mise à jour des clauses plus sur les autres critères}
  GenInf.Plus := SQL;
  GenSup.Plus := SQL;
  ColInv.Plus := SQL;

  NatureAuxi.Plus := '';
  {Si pas de filtre sur la nature des généraux}
  if NatureGene.NbSelected = 0 then begin
    NatureAuxi.Enabled := True;
    AuxInf.Enabled := True;
    AuxSup.Enabled := True;
  end

  {S'il n'y a que les TIC et/ou TID}
  else if (Pos('COS', NatureGen) = 0) and (Pos('COC', NatureGen) = 0) and
     (Pos('COC', NatureGen) = 0) and (Pos('COC', NatureGen) = 0) and
    ((Pos('TIC', NatureGen) > 0) or (Pos('TID', NatureGen) > 0)) then begin
    AuxInf.Enabled := False;
    AuxSup.Enabled := False;
    NatureAuxi.Enabled := False;
  end

  else begin
    NatureAuxi.Enabled := True;
    AuxInf.Enabled := True;
    AuxSup.Enabled := True;

    Lib := '';
    {Exclusion de la Nature des auxiliaires en fonction de la nature des généraux}
    { FQ 18571 BVE 19.04.07 }
    if NatureGene.Tous then
       Lib := ''
    else
    begin
       if Pos('COS', NatureGen) = 0 then Lib := ' AND CO_CODE <> "SAL" ';
       if Pos('COD', NatureGen) = 0 then Lib := Lib + 'AND CO_CODE <> "DIV" ';
       if Pos('COF', NatureGen) = 0 then Lib := Lib + 'AND CO_LIBRE <> "F" ';
       if Pos('COC', NatureGen) = 0 then Lib := Lib + 'AND CO_LIBRE <> "C" ';
    end;             
    { END FQ 18571 }

    NatureAuxi.Plus := Lib;
  end;
  {Réinitialise la combo des natures d'auxiliaires et lance NatureAuxiOnChange}
  if (NatureAuxi.Items.Count = 1) and NatureAuxi.Enabled then SetControlText('T_NATUREAUXI', NatureAuxi.Values[0])
                                                         else SetControlText('T_NATUREAUXI', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHEANCIER.OnLoad;
{---------------------------------------------------------------------------------------}
var
  St,St1,
  St2,St3,
  St4,St5,
  S,S1,
  TriLibelle : string;
  i, j  : Integer;
  SQL   : string;
  Lib   : string;
  Order : string;
  Where : string;
begin
  inherited;
  S  := '';
  S1 := '';
  j  := 0;

  St  := GetControlText('TRIPAR');
  St1 := GetControlText('RUPTURE1');
  St2 := GetControlText('RUPTURE2');

  if GetControlText('TRILIBELLE') = 'X' then TriLibelle := 'T_LIBELLE'
                                        else TriLibelle := 'E_AUXILIAIRE';

  while St1 <> '' do
    For i := 0 to 3 do begin
      St3 := ReadTokenSt(St1);
      St4 := ReadTokenSt(St2);
      if (St3 <> '-') and (St3 > '#') and (St3 <> '') then begin
        if (St3 <> '*') then
          if not TCheckBox(GetControl('CPTASSOCIES')).Checked then begin
            if j < 1 then S := S + ' AND (E_TABLE'+ IntToStr(i) + ' >= "' + St3 + '" AND E_TABLE' + IntToStr(i) + ' <= "' + St4 + '") '
                     else S := S + ' OR (E_TABLE' + IntToStr(i) + ' >= "' + St3 + '" AND E_TABLE' + IntToStr(i) + ' <= "' + St4 + '") ';
            j:=j+1 ;
          end
          else
            S := S + ' AND (E_TABLE' + IntToStr(i) + ' >= "' + St3 + '" AND E_TABLE' + IntToStr(i) + ' <= "' + St4 + '") ';

        if (St3='*') then
          if TCheckBox(GetControl('CPTASSOCIES')).Checked then begin
            S := S + ' AND (E_TABLE' + IntToStr(i) + ' <> "") ';
            j := j + 1;
          end;
      end;{if (St3<>'-') and }
    end;{For i :}

  while St <> '' do
    for i := 0 to 3 do begin
      St5 := ReadTokenSt(St);
      if (St5 <> '') then begin
        SetControlText('ORDER' + IntToStr(i), 'E_TABLE' + Copy(St5, 3, 1));
        SetControlText('E0'    + IntToStr(i), St5);
        if (St = '') then S1 := S1 + 'E_TABLE' + Copy(St5, 3, 1)
                     else S1 := S1 + 'E_TABLE' + Copy(St5, 3, 1)+',';
      end
      else begin
        SetControlText('E0'    + IntToStr(i), '');
        SetControlText('ORDER' + IntToStr(i), '');
      end ;
    end;

  case SensMouvement.ItemIndex of
    0 : S := S + ' AND ' + GetControlText('CREDIT') + ' <> 0 ';
    1 : S := S + ' AND ' + GetControlText('DEBIT' ) + ' <> 0 ';
  end;

  {FQ 17263 : Gestion des comptes d'exclusion}
  St := GetControlText('COMPTESEXCLUS');
  if St <> '' then begin
    S := S + 'AND E_GENERAL NOT IN (';
    while St <> '' do begin
      St1 := ReadTokenSt(St);
      S := S + '"' + St1 + '",';
    end;
    S[Length(S)] := ')';
  end;

  Order := '';
  if S1 <> '' then Order := S1 + ',';
  Order := Order + 'E_DATEECHEANCE,E_MODEPAIE,'+TriLibelle+',E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE';

  SetControlText('XX_WHERE', S);
  SetControlText('XX_ORDERBY', Order);

  if GetControlText('G_NATUREGENE' ) = '' then SetControlText('G_NATUREGENE', 'COF;COC;COS;COD;TIC;TID;');
  if GetControlText('E_QUALIFPIECE') = '' then SetControlText('E_QUALIFPIECE', 'N;S;U;');
  TraductionTHMultiValComboBox(THMultiValComboBox(GetControl('E_QUALIFPIECE')), SQL, Lib, 'E_QUALIFPIECE');
  SetControlText('LIBQUALIF', Lib);
  

  // FQ 18007 BVE 18.04.07
  // On supprime le order
  Where := Copy(RecupWhereCritere(TPageControl(GetControl('PAGES'))),1,Pos('ORDER',RecupWhereCritere(TPageControl(GetControl('PAGES'))))-1);
  SetControlText('WHERESQL', Where);
  // END FQ 18007
end;

initialization
  RegisterClasses([TOF_ECHEANCIER]) ;

end.

