{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/09/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPTPAYEURGLAGE ()
Mots clefs ... : TOF;CPTPAYEURGLAGE
*****************************************************************}
Unit CPTPAYEURGLAGE_TOF ;

Interface

uses Classes, StdCtrls,
{$IFDEF EAGLCLIENT}
     eQRS1,MaineAGL, UTob,
{$ELSE}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     QRS1,FE_Main,
{$ENDIF}
     comctrls,UTof, HCtrls,Ent1,  TofMeth, HTB97, spin,
     Filtre,
     SysUtils,
     ParamSoc,		//GetParamSocSecur YMO
     HQry; {RecupWhereCritere}

Type
  TOF_CPTPAYEURGLAGE = Class (TOF_Meth)
  private
    EtatType	    : String;
    NbPeriodes	    : Integer;
    Periodicite     : THValComboBox ;
    Devise          : THValComboBox ;
    NatureAuxi      : THValComboBox ;
    DevPivot        : THEdit ;
    Ecart           : THRadioGroup ;
    Affichage       : THRadioGroup ;
    AffMontants     : THRadioGroup ;
    Impression      : THRadioGroup ;
    NbJEcart        : TSpinEdit ;
    DateArrete	    : THEdit;
    {FP Pas de situation EnSituation	    : TCheckBox;}
    ConditionSQL    : THEdit;
    Gene1,Gene2	    : THEdit;
    QualifPiece     : THMultiValcomboBox;
    AvecQualifPiece : THEdit;
    {FP}
    ModePaiement    : THMultiValcomboBox ;
    SortSens        : THValComboBox ;
    AvecSens        : THEdit;
    {FP}
    BalPrev, BalRet : TRadioButton;
    ChoixTypeTaLi   : THRadioGroup;
    procedure NatureAuxiOnChange (Sender : TObject) ;
    procedure PeriodiciteOnChange (Sender : TObject) ;
    procedure DateArreteOnChange (Sender : TObject) ;
    procedure EcartOnChange (Sender : TObject) ;
    procedure NbJEcartOnChange(Sender : TObject) ;
    procedure EtablissementOnChange (Sender : TObject) ;
    procedure SortSensOnChange (Sender : TObject) ;
    procedure ModePaiementOnChange (Sender : TObject) ;
    procedure DeviseOnChange (Sender : TObject) ;
    procedure Gene1OnChange (Sender : TObject) ;
    procedure Gene2OnChange (Sender : TObject) ;
    procedure Rupture1OnDblClick (Sender : TObject) ;
    procedure TriParOnDblClick (Sender : TObject) ;
    procedure AffichageOnClick (Sender : TObject) ;
    procedure ImpressionOnClick (Sender : TObject );
    procedure TriParOnChange (Sender : TObject) ;
    function  GetMinMaxCompte(stTable, stCol, stDefaut : String; TiersPayeur: Boolean = False) : String;
    function  GetTabletteColl(stNatureTiers : String) : String;
    function  GetTabletteTiers(stNatureTiers : String) : String;
    function  GetTabletteTiersPayeur(stNatureTiers: String): String;
    function  GetWhereColl(stNatureTiers : String): String;
    procedure GenererConditionSQL;
    procedure ParametreEtatType;
    procedure	CalculPeriodesAvecPeriodicite;
    procedure	CalculPeriodesAvecNbJour;
    procedure BalTypeOnChange(Sender : TObject);
    procedure UpdateOptions( vBoAvecRupt : Boolean ) ;
    procedure OnClickChoixTaLi ( Sender : TObject );
    procedure AuxiElipsisClick ( Sender : TObject );
  public
    procedure OnArgument(Arguments : string); override ;
    procedure OnLoad; override ;
    procedure OnNew; override ;
    procedure OnUpdate; override ;
  end ;

  procedure CPLanceFiche_CPTPAYEURGLAGE;
  procedure CPLanceFiche_CPTPAYEURGLVENTIL;

Implementation

uses HEnt1, LicUtil,HMsgBox, UTofMulParamGen,
     TofTLTiersTri, TofTLTiersRupt;{JP 25/08/04 : FQ 13401, 13452, 13671}

Type TTabDate8 = Array[1..8] of TDateTime ;

procedure CPLanceFiche_CPTPAYEURGLAGE;
begin
  AGLLanceFiche('CP','CPTPAYEURGLAGE','','','AGEE');
end;

procedure CPLanceFiche_CPTPAYEURGLVENTIL;
begin
  AGLLanceFiche('CP','CPTPAYEURGLAGE','','','VENTIL');
end;

PROCEDURE TOF_CPTPAYEURGLAGE.OnUpdate ;
VAR St,St1,St2,St3,St4,St5,S,S1,TriLibelle,Order,stTL : string ;
    i : integer;
    stIndice : string;
BEGIN
  S:='';
  S1:='';
  TriLibelle:='T_AUXILIAIRE' ;
  St:=GetControlText('TRIPAR') ;
  St1:=GetControlText('RUPTURE1') ;
  St2:=GetControlText('RUPTURE2') ;
  if GetControlText('TRILIBELLE')='X'
    then TriLibelle:='T_LIBELLE'
    else TriLibelle:='T_AUXILIAIRE' ;

  // Construction condition sur ruptures tables libres
  // Codes utilisés pour la gestion des ruptures sur tables libres :
  //    # = non gérée
  //    - = gérée mais non sélectionnée
  //    * = tous les enregistrements de la table (sans fourchette de valeurs)
  //  sinon code de valeur début dans St3, fin dans St4
  if St1<>'' then
    For i:=0 to 9 do
      BEGIN
      St3:=ReadTokenSt(St1) ;
      St4:=ReadTokenSt(St2) ;
      if (St3<>'-') and (St3>'#') and (St3<>'') then // unniquement les tables sélectionnées
        BEGIN
        if (St3<>'*') then // Si on a une fourchette de valeurs
          begin
            if ChoixTypeTaLi.Value = 'AUX' then
              stTL := '( T_TABLE'+IntToStr(i)+'>="'+St3+'" AND T_TABLE'+IntToStr(i)+'<="'+St4+'" )'
            else if ChoixTypeTaLi.Value = 'CLI' then
              stTL := '( YTC_TABLELIBRETIERS'+IntToHex(i+1,1)+'>="'+St3+'" AND YTC_TABLELIBRETIERS'+IntToHex(i+1,1)+'<="'+St4+'" )'
            else if ChoixTypeTaLi.Value = 'SCL' then
              stTL := '( YTC_RESSOURCE'+IntToHex(i+1,1)+'>="'+St3+'" AND YTC_RESSOURCE'+IntToHex(i+1,1)+'<="'+St4+'" )';
          if not TCheckBox(GetControl('CPTASSOCIES')).Checked then
          begin
            if ChoixTypeTaLi.Value = 'AUX' then
              stTL := '( ' + stTL +' OR ( T_TABLE'+IntToStr(i)+'="" ) ) '
            else if ChoixTypeTaLi.Value = 'CLI' then
              stTL := '( ' + stTL +' OR ( YTC_TABLELIBRETIERS'+IntToHex(i+1,1)+'="" ) ) '
            else if ChoixTypeTaLi.Value = 'SCL' then
              stTL := '( ' + stTL +' OR ( YTC_RESSOURCE'+IntToHex(i+1,1)+'="" ) ) ' ;
          end;
          S := S + ' AND ' +  stTL ;
          end
        else // Sans fourchette de valeur
          if TCheckBox(GetControl('CPTASSOCIES')).Checked then
          begin
            if ChoixTypeTaLi.Value = 'AUX' then
              S := S + ' AND (T_TABLE'+IntToStr(i)+'<>"") '
            else if ChoixTypeTaLi.Value = 'CLI' then
              S := S + ' AND (YTC_TABLELIBRETIERS'+IntToHex(i+1,1)+'<>"") '
            else if ChoixTypeTaLi.Value = 'SCL' then
              S := S + ' AND (YTC_RESSOURCE'+IntToHex(i+1,1)+'<>"") ' ;
          end;
        END ;
      END ;

  {
  // conditions auxiliaire
  if (GetControlText('T_AUXILIAIRE')='')  and (GetControlText('T_AUXILIAIRE_')<>'') then
    S:=S+' AND T_AUXILIAIRE<="'+GetControlText('T_AUXILIAIRE_')+'"' ;
  if (GetControlText('T_AUXILIAIRE')<>'') and (GetControlText('T_AUXILIAIRE_')<>'') then
    S:=S+' AND T_AUXILIAIRE>="'+GetControlText('T_AUXILIAIRE')+'" AND T_AUXILIAIRE<="'+GetControlText('T_AUXILIAIRE_')+'"' ;

  {b FP
  // conditions tiers payeur
  if (GetControlText('T_TIERSPAYEUR')='')  and (GetControlText('T_TIERSPAYEUR_')<>'') then
    S:=S+' AND E_TIERSPAYEUR<="'+GetControlText('T_TIERSPAYEUR_')+'"' ;
  if (GetControlText('T_TIERSPAYEUR')<>'') and (GetControlText('T_TIERSPAYEUR_')<>'') then
    S:=S+' AND E_TIERSPAYEUR>="'+GetControlText('T_TIERSPAYEUR')+'" AND E_TIERSPAYEUR<="'+GetControlText('T_TIERSPAYEUR_')+'"' ;

  // Condition sur le mode de paiement
  if (ModePaiement <> nil) and (not ModePaiement.Tous) then
    S:=S+ModePaiement.GetSQLValue;

  inutile avec RecupWhereSQL  }

  // Condition sur le sens
  if (SortSens <> nil) and (SortSens.ItemIndex=0) then
    S:=S+' AND '+GetControlText('CREDIT')+'<>0'
  else if (SortSens <> nil) and (SortSens.ItemIndex=1) then
    S:=S+' AND '+GetControlText('DEBIT')+'<>0';
  {e FP}

  // conditions nature auxi
{  if TComboBox(GetControl('T_NATUREAUXI')).itemindex<>0 then
    S:=S+' AND T_NATUREAUXI="'+GetControlText('T_NATUREAUXI')+'"' ;
  inuyile avec RecupWhereSQL  }


  // Mise en place dans le champ caché
  SetControlText('WHERE',S) ;

  // MAJ pour gestion des ruptures dans l'état
  for i:=0 to 9 do
    begin
    SetControlText('T0'+IntToStr(i),'') ;
    SetControlText('ORDER'+IntToStr(i),'') ;
    end ;
{b FP FQ18231 30/05/2006}
  SetControlText('ORDERA','E_TIERSPAYEUR') ;
  i := 0;
  while (St<>'') and (i < 10) do
    begin
{e FP FQ18231 30/05/2006}
    St5:=ReadTokenSt(St) ;
    if (St5<>'') then
      begin
        stIndice := copy(St5,3,1);
      if ChoixTypeTaLi.Value = 'AUX' then
        SetControlText('ORDER'+IntToStr(i),'T_TABLE'+stIndice)
      else if ChoixTypeTaLi.Value = 'CLI' then SetControlText('ORDER'+IntToStr(i),'YTC_TABLELIBRETIERS'+IntToHex(StrToInt(stIndice),1))
      else if ChoixTypeTaLi.Value = 'SCL' then SetControlText('ORDER'+IntToStr(i),'YTC_RESSOURCE'+IntToHex(StrToInt(stIndice),1));
      SetControlText('T0'+IntToStr(i),St5) ;
      if (St='') then
      begin
        if ChoixTypeTaLi.Value = 'AUX' then
          S1:=S1+'T_TABLE'+copy(St5,3,1)
        else if ChoixTypeTaLi.Value = 'CLI' then S1:=S1+'YTC_TABLELIBRETIERS'+IntToHex(StrToInt(stIndice),1)
        else if ChoixTypeTaLi.Value = 'SCL' then S1:=S1+'YTC_RESSOURCE'+IntToHex(StrToInt(stIndice),1)
      end
      else
      begin
        if ChoixTypeTaLi.Value = 'AUX' then
          S1:=S1+'T_TABLE'+copy(St5,3,1)+','
        else if ChoixTypeTaLi.Value = 'CLI' then S1:=S1+'YTC_TABLELIBRETIERS'+IntToHex(StrToInt(stIndice),1)+','
        else if ChoixTypeTaLi.Value = 'SCL' then S1:=S1+'YTC_RESSOURCE'+IntToHex(StrToInt(stIndice),1)+',';
      end;
    end ;
    inc(i) ;
    end ;
{b FP FQ18231 30/05/2006}
  SetControlText('ORDER10',TriLibelle) ;
{e FP FQ18231 30/05/2006}

  if S1<>''
    then Order:=S1+','+TriLibelle
    else Order:=TriLibelle ;

  Order:='E_TIERSPAYEUR,'+Order+',E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE'; {FP}


  { Mise à jour des requêtes pour récupérer les libellés des données des tables libres }
  for i := 0 to 9 do
  begin
    if ChoixTypeTaLi.Value = 'AUX' then
      SetControlText ('SQL'+IntToStr(i),'select nt_libelle FROM natcpte WHERE nt_typecpte="T0'+IntToStr(i)+'" and nt_nature=')
    else if ChoixTypeTaLi.Value = 'CLI' then SetControlText ('SQL'+IntToStr(i),'select yx_libelle FROM choixext WHERE yx_type="LT'+IntToHex(i+1,1)+'" and yx_code=')
    else if ChoixTypeTaLi.Value = 'SCL' then SetControlText ('SQL'+IntToStr(i),'select ars_libelle FROM ressource WHERE ars_ressource=');
  end;

  // YMO - 09/02/2007 - FQ 19557 Changement des champs T_ sur la fiche nécessaire
  TFQRS1(Ecran).WhereSQL:=StringReplace(TFQRS1(Ecran).WhereSQL, 'T_', 'E_', [rfReplaceAll]) + '|'+Order ;

END;

PROCEDURE TOF_CPTPAYEURGLAGE.OnArgument(Arguments : string);
VAR Rupture1,Rupture2               : THEdit;
    TriPar                          : THEdit;
BEGIN
  inherited ;

  // Mémorisation du type d'état (AGEE ou VENTIL)
  EtatType := Arguments;
  // Gestion du filtre et du nb de périodes suivant type de l'état
  if EtatType = 'VENTIL' then
    begin
    NbPeriodes := 4 ;
    TFQRS1(Ecran).FNomFiltre := 'TTPGLVENTIL' ; {FP Modif du nom du filtre}
    Ecran.HelpContext := 7237000;
    end
  else
    begin
    NbPeriodes := 4 ;                           {FP NbPeriodes := 5}
    TFQRS1(Ecran).FNomFiltre := 'TTPGLAGEE' ;   {FP Modif du nom du filtre}
    Ecran.HelpContext := 7236000;               {FP Modif de l'aide}
    end ;

  NatureAuxi    := THValComboBox(GetControl('T_NATUREAUXI'));
  Periodicite   := THValComboBox(GetControl('PERIODICITE'));
  Devise        := THValComboBox(GetControl('DEVISE'));
  DevPivot      := THEdit(GetControlText('DEVPIVOT'));
  Gene1         := THEdit(GetControl('GENERAL'));
  Gene2         := THEdit(GetControl('GENERAL_'));
  DateArrete    := THEdit(GetControl('DATE1'));
  ConditionSQL  := THEdit(GetControl('CONDITIONSQL'));
  Ecart         := THRadioGroup(GetControl('ECART'));
  NbJEcart      := TSpinEdit(GetControl('NBJECART'));
  Rupture1      := THEdit(GetControl('RUPTURE1'));
  Rupture2      := THEdit(GetControl('RUPTURE2'));
  TriPar        := THEdit(GetControl('TRIPAR'));
  Affichage     := THRadioGroup(GetControl('AFFICHAGE')) ;
  AffMontants   := THRadioGroup(GetControl('AFFMONTANTS')) ;
  Impression    := THRadioGroup(GetControl('TIMPRESSION')) ;
  QualifPiece   := THMultiValComboBox(Getcontrol('QUALIFPIECE'));
  AvecQualifPiece :=THEdit(GetControl('AVECQUALIFPIECE'));
  {FP}
  ModePaiement  := THMultiValcomboBox(GetControl('E_MODEPAIE'));
  SortSens      := THValComboBox(GetControl('SORTESENS'));
  AvecSens      := THEdit(GetControl('AVECSENS'));
  {FP}
  BalPrev       :=TRadioButton(GetControl('BALPREV'));
  BalRet        :=TRadioButton(GetControl('BALRET'));
  ChoixTypeTaLi :=THRadioGroup(GetControl('CHOIXTYPETALI'));
  if Devise<>nil        then BEGIN Devise.OnChange:=DeviseOnChange;Devise.itemindex:=0;Devise.OnChange(NIL) ; END ;
  if ComboEtab <> nil   then BEGIN ComboEtab.OnChange:=EtablissementOnChange;ComboEtab.OnChange(NIL) ; END ;
  if Periodicite<>nil   then BEGIN Periodicite.OnChange:=PeriodiciteOnChange;Periodicite.itemindex:=0 ; END ;
  if DateArrete<>nil    then BEGIN DateArrete.OnChange:=DateArreteOnChange;SetControlText('DATE1',DateToStr(V_PGI.DateEntree)); END ;
  if Ecart<>nil         then BEGIN Ecart.OnClick:=EcartOnChange; Ecart.itemindex:=0;EcartOnChange(NIL) ; END ;
  if NbJEcart<>nil      then BEGIN NbJEcart.OnChange:=NbJEcartOnChange;NbJEcart.text:='30' ; END ;
  if NatureAuxi<>nil    then BEGIN NatureAuxi.OnChange:=NatureAuxiOnChange;NatureAuxi.itemindex:=0; NatureAuxiOnChange(nil); END ;
  if ChoixTypeTaLi <> nil then BEGIN ChoixTypeTaLi.OnClick := OnClickChoixTaLi; ChoixTypeTali.ItemIndex := 0; End;
  {FP}
  if SortSens<>nil      then BEGIN SortSens.OnChange:=SortSensOnChange;SortSens.itemindex:=2;SortSens.OnChange(NIL) ; END ;
  if ModePaiement<>nil  then BEGIN ModePaiement.OnChange:=ModePaiementOnChange;ModePaiement.Value:=TraduireMemoire('<<Tous>>') ; END ;
  {FP}
  if QualifPiece<>nil		then QualifPiece.value := 'N';
  if Affichage<>nil then
    BEGIN
    Affichage.OnClick:=AffichageOnClick;
    Affichage.itemindex:=0;
    SetControlText('DEBIT','E_DEBIT') ;
    SetControlText('CREDIT','E_CREDIT') ;
    SetControlText('COUVERTURE','E_COUVERTURE') ;
    END ;
  if AffMontants<>nil then
    BEGIN
    AffMontants.itemindex:=0;
    SetControlText('AFFMONTANT','NORMAL') ;
    END ;
  if Impression<>nil then
    BEGIN
    Impression.OnClick:=ImpressionOnClick;
    Impression.itemindex:=0;
    SetControlText('IMPRESSION','DECI');
    END ;
  if (Gene1<>nil) and (Gene2<>nil) then
    BEGIN
    Gene1.OnChange:=Gene1OnChange;
    Gene2.OnChange:=Gene2OnChange;
    END ;
  if Rupture1<>nil then Rupture1.OnDblClick:=Rupture1OnDblClick;
  if Rupture2<>nil then Rupture2.OnDblClick:=Rupture1OnDblClick;
  if TriPar<>nil   then
    BEGIN
    TriPar.OnChange:=TriParOnChange;
    TriPar.OnDblClick:=TriParOnDblClick;
    SetControlEnabled('SAUTPAGE',False);
    SetControlEnabled('SURRUPTURE',False);
    END ;
  BalPrev.OnClick := BalTypeOnChange;
  BalRet.OnClick := BalTypeOnChange;
  //=====> Param selon type de balance
  ParametreEtatType;
  // Init options des ruptures
  UpdateOptions(false);
  ChoixTypeTali.Visible := TL_TIERSCOMPL_Actif;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    THEdit(GetControl('T_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;
    THEdit(GetControl('T_AUXILIAIRE_', true)).OnElipsisClick:=AuxiElipsisClick;
  end;

END;

procedure TOF_CPTPAYEURGLAGE.OnLoad;
var stTypeEcr : String;
begin
  inherited;
  // Auto-remplissage si comptes auxiliaires non renseignés
  if (GetControlText('T_AUXILIAIRE')='') then
    SetControlText('T_AUXILIAIRE',GetMinMaxCompte('TIERS', 'MIN(T_AUXILIAIRE)','0'));
  if (GetControlText('T_AUXILIAIRE_')='') then
    SetControlText('T_AUXILIAIRE_',GetMinMaxCompte('TIERS', 'MAX(T_AUXILIAIRE)','ZZZZZZZZZZZZZZZZZ'));

  // Auto-remplissage si tiers payeur non renseignés
  if (GetControlText('T_TIERSPAYEUR')='') then
    SetControlText('T_TIERSPAYEUR',GetMinMaxCompte('TIERS', 'MIN(T_AUXILIAIRE)','0', True));
  if (GetControlText('T_TIERSPAYEUR_')='') then
    SetControlText('T_TIERSPAYEUR_',GetMinMaxCompte('TIERS', 'MAX(T_AUXILIAIRE)','ZZZZZZZZZZZZZZZZZ', True));

  // Auto-remplissage si comptes collectifs non renseignés
  if (Gene1.Text='') then begin
    Gene1.Text := GetMinMaxCompte('GENERAUX', 'MIN(G_GENERAL)','0');
		SetControlText('COLLECTIF1',GetControlText('GENERAL'));
    end;
  if (Gene2.Text='') then begin
    Gene2.Text := GetMinMaxCompte('GENERAUX', 'MAX(G_GENERAL)','ZZZZZZZZZZZZZZZZZ');
		SetControlText('COLLECTIF2',GetControlText('GENERAL_'));
    end;
  // Affichage des montants
  if AffMontants.ItemIndex = 1 then
    SetControlText('AFFMONTANT','SIGNES')
  else
    SetControlText('AFFMONTANT','NORMAL');

  // libellés pour les types d'écritures...
  AvecQualifPiece.Text := '';
  stTypeEcr := QualifPiece.Value;
  if stTypeEcr <> '' then
    while stTypeEcr <> '' do
      AvecQualifPiece.Text := AvecQualifPiece.Text
      	+ RechDom(QualifPiece.DataType,ReadTokenSt(stTypeEcr),False) + ' '
  else
    AvecQualifPiece.Text := '<<Tous>>';

  // Si exoV8 renseigné, alors on ne prend que les ecritures post début d'exoV8
  if (VH^.ExoV8.Code <> '') then
    SetControlText('DateExoV8', DateTimeToStr(VH^.ExoV8.Deb))  {FP FQ17987 30/05/2006}
  else
    SetControlText('DateExoV8', DateTimeToStr(iDate1900));     {FP FQ17987 30/05/2006}

  // Remplir les conditions SQL
  GenererConditionSQL;
end;

//==============================================================================

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 02/06/2005
Modifié le ... : 02/06/2005
Description .. : Ajout du paramètre de type de table libre pour appel de la 
Suite ........ : fonction de paramétrage des tables libres
Suite ........ : 
Mots clefs ... :
*****************************************************************}
PROCEDURE TOF_CPTPAYEURGLAGE.Rupture1OnDblClick (Sender : TObject);
VAR    St,St1,St2,St3,Arg : String ;
       i : integer ;
       OkAssoc : integer ;
BEGIN
  OkAssoc := 0;
  Arg     := GetControlText('RUPTURE1')+'|'+GetControlText('RUPTURE2');
  if GetControlVisible('CHOIXTYPETALI') then
    Arg := Arg + '|' + GetControlText('CHOIXTYPETALI')
  else Arg:= Arg + '|AUX';
  St      := CPLanceFiche_TLTIERSRUPT(Arg); {JP 25/08/04 : FQ 13401, 13452, 13671}
  St1     := ReadTokenPipe(St,'|');
  SetControlText('RUPTURE1',St1);
  SetControlText('RUPTURE2',St);
  for i:=0 to 9 do
    BEGIN
    St2:=ReadTokenSt(St1);
    if (St2<>'#') and (St2<>'-') and (St2<>'') then
      BEGIN
        if ChoixTypeTaLi.Value = 'AUX' then
                St3:=St3+'T0'+IntToStr(i)+';'
        else if ChoixTypeTaLi.Value = 'CLI' then
              St3:=St3+'LT'+IntToHex(i+1,1)+';'
        else St3 :=St3+'AR'+IntToHex(i+1,1)+';';
      OkAssoc:=OkAssoc+1 ;
      END ;
    SetControlText('TRIPAR',St3);
    END ;

  // MAJ interface si aucune rupture sélectionnée
  UpdateOptions( OkAssoc > 0 ) ;
  if OkAssoc = 0 then
    begin
    SetControlText('RUPTURE1', '');
    SetControlText('RUPTURE2', '');
    SetControlText('TRIPAR',   '');
    end ;
END;

PROCEDURE TOF_CPTPAYEURGLAGE.TriParOnChange (Sender : TObject);
BEGIN
if GetControlText('TRIPAR')='' then
  BEGIN
  SetControlChecked('SAUTPAGE',False);
  SetControlChecked('SURRUPTURE',False);
  END else
  BEGIN
  SetControlEnabled('SAUTPAGE',True);
  SetControlEnabled('SURRUPTURE',True);
  end;
END;

PROCEDURE TOF_CPTPAYEURGLAGE.TriParOnDblClick (Sender : TObject);
VAR    St,Arg:String;
BEGIN
Arg := GetControlText('TRIPAR');
St  := CPLanceFiche_TLTIERSTRI(Arg); {JP 25/08/04 : FQ 13401, 13452, 13671}
SetControlText('TRIPAR',St);
END;

PROCEDURE TOF_CPTPAYEURGLAGE.Gene1OnChange (Sender : TObject);
BEGIN
SetControlText('COLLECTIF1',GetControlText('GENERAL'));
END;

PROCEDURE TOF_CPTPAYEURGLAGE.Gene2OnChange (Sender : TObject);
BEGIN
SetControlText('COLLECTIF2',GetControlText('GENERAL_'));
END;

PROCEDURE TOF_CPTPAYEURGLAGE.ImpressionOnClick (Sender : TObject);
BEGIN
case Impression.ItemIndex of
  0 : SetControlText('IMPRESSION','DECI');
  1 : SetControlText('IMPRESSION','KILO');
  2 : SetControlText('IMPRESSION','MEGA');
  3 : SetControlText('IMPRESSION','ENTI');
  end;
END;

PROCEDURE TOF_CPTPAYEURGLAGE.AffichageOnClick (Sender : TObject) ;
BEGIN
case Affichage.ItemIndex of
  0 : BEGIN SetControlText('DEBIT','E_DEBIT')     ; SetControlText('CREDIT','E_CREDIT')     ; SetControlText('COUVERTURE','E_COUVERTURE') ; END ;
  1 : BEGIN SetControlText('DEBIT','E_DEBITDEV')  ; SetControlText('CREDIT','E_CREDITDEV')  ; SetControlText('COUVERTURE','E_COUVERTUREDEV') ; END ;
  end;
END;

PROCEDURE TOF_CPTPAYEURGLAGE.DeviseOnChange (Sender : TObject);
BEGIN
  if Devise.ItemIndex=0 then
	  begin
  	SetControlText('DEVPIVOT','%');
    Affichage.ItemIndex := 0;
    Affichage.Enabled := False;
	  end
  else
  	begin
	  SetControlText('DEVPIVOT',GetControlText('DEVISE'));
    Affichage.Enabled := True;
  	end;
END ;

PROCEDURE TOF_CPTPAYEURGLAGE.EtablissementOnChange (Sender : TObject);
BEGIN
if ComboEtab.ItemIndex=0 then SetControlText('ETAB','%')
                         else SetControlText('ETAB',GetControlText('ETABLISSEMENT'));
END;

PROCEDURE TOF_CPTPAYEURGLAGE.SortSensOnChange (Sender : TObject) ;
BEGIN
if SortSens.ItemIndex=0 then SetControlText('AVECSENS','%') else SetControlText('AVECSENS',GetControlText('SORTESENS'));
END;

PROCEDURE TOF_CPTPAYEURGLAGE.ModePaiementOnChange (Sender : TObject) ;
BEGIN
if ModePaiement.Tous then SetControlText('MPAIE',TraduireMemoire('<<Tous>>'))
else SetControlText('MPAIE',ModePaiement.Text);
END;

PROCEDURE TOF_CPTPAYEURGLAGE.NatureAuxiOnChange (Sender : TObject) ;
Var CGen, CAux : string ;
BEGIN
inherited ;
SetControlText('T_AUXILIAIRE','');
SetControlText('T_AUXILIAIRE_','');
{b FP}
SetControlText('T_TIERSPAYEUR','');
SetControlText('T_TIERSPAYEUR_','');
{e FP}
SetControlText('GENERAL','');
SetControlText('GENERAL_','');
SetControlText('NATUREAUX',GetControlText('T_NATUREAUXI'));

CAux := GetTabletteTiers(NatureAuxi.value);
SetControlProperty('T_AUXILIAIRE','DATATYPE',CAux);
SetControlProperty('T_AUXILIAIRE_','DATATYPE',CAux);
{b FP}
CAux := GetTabletteTiersPayeur(NatureAuxi.value);
SetControlProperty('T_TIERSPAYEUR','DATATYPE',CAux);
SetControlProperty('T_TIERSPAYEUR_','DATATYPE',CAux);
{e FP}
CGen := GetTabletteColl(NatureAuxi.value);
SetControlProperty('GENERAL','DATATYPE',CGen);
SetControlProperty('GENERAL_','DATATYPE',CGen);
END ;

PROCEDURE PeriodiciteChange(Iperiodicite : Integer ; FP8,FP1 : THEdit ; var TabD : TTabDate8) ;
VAR Choix,i : Integer ;
    a,m,j, JMax : Word ;
    DAT : TDATETIME ;
BEGIN
//Choix:=FPeriodicite.ItemIndex ;
Choix:=IPeriodicite ;
If Choix<=-1 Then Choix:=0 ;
   TabD[8]:=StrToDate(FP8.text) ;       { Calcul à partir de la date d'arrêtée }
   Case IPeriodicite of
     0,1,2,3,4,5 : for i:=5 downto 1 Do { Pour chaque Fourchette de dates( en partant de la derniére) , en mensuel }
                       BEGIN
                       DAT:=PlusMois(TabD[i+5],-(Choix+1)) ;
                       DecodeDate(DAT,a,m,j) ;
                       JMax:=StrToInt(FormatDateTime('d',FinDeMois(EncodeDate(a,m,1)))) ;
                       TabD[i]:=PlusMois(TabD[i+5],-(Choix+1))+(JMax-J)+1 ;
                       if i>1 then TabD[i+4]:=TabD[i]-1 ;
                       END ;
     6           : for i:=5 downto 1 Do { Pour chaque Fourchette de dates( en partant de la derniére) , en Quinzaine }
                       BEGIN
                       DAT:=TabD[i+5] ;
                       DecodeDate(DAT,a,m,j) ;
                       JMax:=StrToInt(FormatDateTime('d',FinDeMois(EncodeDate(a,m,1)))) ;
                       If J<=15 then
                          BEGIN         { Date départ = (Date d'arrivée - 15 jours) + 1 jour, si date avant le 15 du mois }
                          TabD[i]:=TabD[i+5]-(15  )+1 ;
                          END Else
                          BEGIN         { Date départ = (Date d'arrivée - (Nb jours Max du mois - 15 jours) ) + 1 jour, si date aprés le 15 du mois }
                          TabD[i]:=TabD[i+5]-(JMax-15)+1 ;
                          END ;
                       if i>1 then TabD[i+4]:=TabD[i]-1 ;
                       END ;
     7           : for i:=5 downto 1 Do { Pour chaque Fourchette de dates( en partant de la derniére) , en Hebdo }
                       BEGIN            { Date départ = (Date d'arrivée - (7 jours+ 1 jours) )  }
                       TabD[i]:=TabD[i+5]-(7)+1 ;
                       if i>1 then TabD[i+4]:=TabD[i]-1 ;
                       END ;
     End ;
END;

PROCEDURE TOF_CPTPAYEURGLAGE.PeriodiciteOnChange(Sender : TObject);
BEGIN
inherited ;
CalculPeriodesAvecPeriodicite;
END;

PROCEDURE TOF_CPTPAYEURGLAGE.DateArreteOnChange(Sender : TObject);
BEGIN
inherited ;
if not IsValidDate(GetControlText('DATE1')) then exit;
if (EtatType = 'VENTIL') and (TRadioButton(GetControl('BALPREV')).Checked)
	then SetControlText('PERIODE1',GetControlText('DATE1'))
	else SetControlText('PERIODE'+IntToStr(NbPeriodes+4),GetControlText('DATE1'));
EcartOnChange(nil);
END;

PROCEDURE TOF_CPTPAYEURGLAGE.EcartOnChange(Sender : TObject);
BEGIN
inherited ;
SetControlVisible('PERIODES',Ecart.ItemIndex=1);
SetControlEnabled('NBJECART',Ecart.ItemIndex=0);
if Ecart.ItemIndex=0 then NbJEcartOnChange(nil)
                     else PeriodiciteOnChange(nil);
END;

PROCEDURE NbJEcartCalculDate(FP8 : THEdit ; NbjEcart : integer ; var TabD : TTabDate8) ;
VAR i : Integer ;
BEGIN
TabD[8]:=StrToDate(FP8.Text) ;       { Calcul à partir de la date d'arrêtée }
For i:=5 downto 1 do
  BEGIN
    TabD[i]:=TabD[i+5]-(NbJEcart-1);
    if i>1 then TabD[i+4]:=TabD[i]-1;
  end;
END;

PROCEDURE TOF_CPTPAYEURGLAGE.NbJEcartOnChange(Sender : TObject);
BEGIN
inherited;
if GetControlText('NBJECART')='' then exit;
CalculPeriodesAvecNbJour;
END ;

function TOF_CPTPAYEURGLAGE.GetMinMaxCompte(stTable, stCol, stDefaut : String; TiersPayeur: Boolean = False) : String;
var
	Q : TQuery;
	stWhere : String;
begin
  // Condition dépend de la nature auxiliaire
  stWhere := '';
  if stTable = 'GENERAUX' then
    if NatureAuxi.Value <> '' then stWhere := GetWhereColl(NatureAuxi.Value)
    else stWhere := ' WHERE G_COLLECTIF="X"'
  else begin
    {b FP}
    if NatureAuxi.Value <> '' then
      stWhere := ' T_NATUREAUXI="' + NatureAuxi.Value + '"';
    if TiersPayeur then begin
      if stWhere <> '' then
        stWhere := stWhere+' AND';
      stWhere := stWhere+' T_ISPAYEUR = "X"';
      end;
    if stWhere <> '' then
      stWhere := ' WHERE '+stWhere;
    {e FP}
    end;
  // Requête
  Q := OpenSQL('SELECT ' + stCol +' CODE FROM ' + stTable + stWhere,True);
  if not Q.Eof then
    Result := Q.FindField('CODE').asString
  else
    Result := stDefaut;
  Ferme(Q);
end;

function TOF_CPTPAYEURGLAGE.GetTabletteColl(stNatureTiers : String): String;
begin
  if stNatureTiers='CLI' then result := 'TZGCOLLCLIENT' 		else
	  if stNatureTiers='FOU' then result := 'TZGCOLLFOURN' 			else
  		if stNatureTiers='SAL' then result := 'TZGCOLLSALARIE' 		else
			  if stNatureTiers='DIV' then result := 'TZGCOLLDIVERS' 		else
				  if stNatureTiers='AUC' then result := 'TZGCOLLTOUTDEBIT' 	else
					  if stNatureTiers='AUD' then result := 'TZGCOLLTOUTCREDIT' else
           	  result := 'TZGCOLLECTIF';
end;

function TOF_CPTPAYEURGLAGE.GetTabletteTiers(stNatureTiers: String): String;
begin
  if stNatureTiers='CLI' then
    result := 'TZTCLIENT'
  else if stNatureTiers='FOU' then
    result := 'TZTFOURN'
  else if stNatureTiers='SAL' then
    result := 'TZTSALARIE'
  else if stNatureTiers='DIV' then
    result := 'TZTDIVERS'
  else if stNatureTiers='AUC' then
    result := 'TZTTOUTDEBIT'
  else if stNatureTiers='AUD' then
    result := 'TZTTOUTCREDIT'
  else
    result := 'TZTTOUS';
end;

function TOF_CPTPAYEURGLAGE.GetTabletteTiersPayeur(stNatureTiers: String): String;
begin
  if stNatureTiers='CLI' then
    result := 'TZTPAYEURCLI'
  else if stNatureTiers='FOU' then
    result := 'TZTPAYEURFOU'
  else if stNatureTiers='SAL' then
    result := 'TZTPAYEUR'
  else if stNatureTiers='DIV' then
    result := 'TZTPAYEURCLI'
  else if stNatureTiers='AUC' then
    result := 'TZTPAYEURFOU'
  else if stNatureTiers='AUD' then
    result := 'TZTPAYEURCLI'
  else
    result := 'TZTPAYEUR';
end;

function TOF_CPTPAYEURGLAGE.GetWhereColl(stNatureTiers : String): String;
begin
  result := ' WHERE G_COLLECTIF="X"';
  if stNatureTiers='CLI' then
    result := result + ' AND G_NATUREGENE="COC"'
  else if stNatureTiers='FOU' then
    result := result + ' AND G_NATUREGENE="COF"'
  else if stNatureTiers='SAL' then
    result := result + ' AND G_NATUREGENE="COS"'
  else if stNatureTiers='DIV' then
    result := result + ' AND G_NATUREGENE="COD"'
  else if stNatureTiers='AUC' then
    result := result + ' AND (G_NATUREGENE="COC" OR G_NATUREGENE="COD")'
  else if stNatureTiers='AUD' then
    result := result + ' AND (G_NATUREGENE="COF" OR G_NATUREGENE="COD")';
end;

procedure TOF_CPTPAYEURGLAGE.GenererConditionSQL;
  {b FP FQ18231 30/05/2006}
  function WhereCommun(Prefix: String): String;
  var stWhereQP, stTypeEcr : String;
  begin
    // Condition sur les comptes généraux
    Result := ' AND '+Prefix+'E_GENERAL>="' + GetControlText('COLLECTIF1') + '"'+
              ' AND '+Prefix+'E_GENERAL<="' + GetControlText('COLLECTIF2') + '"';
    // Condition sur le type de piece
    Result := Result
      + ' AND '+Prefix+'E_QUALIFPIECE<>"C" AND '+Prefix+'E_ECRANOUVEAU<>"CLO"'
      + ' AND '+Prefix+'E_ECRANOUVEAU<>"OAN"' ;
    stTypeEcr := QualifPiece.Value;
    if stTypeEcr <> '' then
      begin
      stWhereQP := '';
      while stTypeEcr <> '' do
        if stWhereQP = ''then
          stWhereQP := ' AND '+Prefix+'E_QUALIFPIECE IN ("' + ReadTokenSt(stTypeEcr) + '"'
        else stWhereQP := stWhereQP + ',"' + ReadTokenSt(stTypeEcr) + '"';
      Result := Result + stWhereQP + ')';
      end;

    {b FP}
    Result := Result +
         ' AND '+Prefix+'E_ECHE="X"' +
         ' AND (NOT ('+Prefix+'E_ETATLETTRAGE="PL" AND (('+Prefix+'E_DEBIT+'+Prefix+'E_CREDIT)='+Prefix+'E_COUVERTURE)))';
    {e FP}

    // Condition sur l'etablissement
    if ComboEtab.itemIndex > 0 then
      Result := Result + ' AND '+Prefix+'E_ETABLISSEMENT = "' + ComboEtab.value + '"';
    // Condition sur la devise
    if Devise.itemIndex > 0 then
      Result := Result + ' AND '+Prefix+'E_DEVISE = "' + Devise.value + '"';
  end;
begin
  //=======================
  ConditionSQL.Text := WhereCommun('') +
    ' AND (E_QUALIFORIGINE <> "TP" or E_QUALIFORIGINE="")' +  {FP FQ17987 30/05/2006}
    ' AND E_ETATLETTRAGE="TL"';
  SetControlText('CONDITIONSQL_SOUSREQUETE', WhereCommun('E1.')+
    ' AND E1.E_QUALIFORIGINE = "TP"' +
    ' AND E1.E_ETATLETTRAGE<>"TL" AND E1.E_ETATLETTRAGE<>"RI"');
  {e FP FQ18231 30/05/2006}
end;

procedure TOF_CPTPAYEURGLAGE.ParametreEtatType;
var stTitre : String;
begin
  if EtatType = 'VENTIL' then begin
    // Modif titre QRS1
    stTitre := 'Grand livre ventilé par tiers payeur';  {FP Modif du titre}
    TFQRS1(Ecran).Caption := stTitre;
    UpdateCaption(Ecran);
    //Affiche du type de balance (Prevision / Retard)
    SetControlVisible('TBALTYPE',True);
    SetControlVisible('BALPREV',True);
    SetControlVisible('BALRET',True);
    SetControlProperty('BALRET','checked',False);
    SetControlProperty('BALPREV','checked',True);
    // Changement de la nature de l'état
    TFQRS1(Ecran).NatureEtat := 'TFA';                {FP}
    TFQRS1(Ecran).CodeEtat   := 'TGV';                {FP}
    end
  else begin
    // Modif titre QRS1
    stTitre := 'Grand livre âgé par tiers payeur';  {FP Modif du titre}
    TFQRS1(Ecran).Caption := stTitre;
    UpdateCaption(Ecran);
    //Cache le type de balance ventilée (Prevision / Retard)
    SetControlVisible('TBALTYPE',False);
    SetControlVisible('BALPREV',False);
    SetControlVisible('BALRET',False);
    end;
  Ecran.Caption :=TraduireMemoire(Ecran.Caption) ;
  UpdateCaption(Ecran);
end;

procedure TOF_CPTPAYEURGLAGE.CalculPeriodesAvecNbJour;
var
	i, NbJours : Integer ;
  Prevision : Boolean;
	TabD : TTabDate8;
begin
	// Initalisation et paramétrage suivant type état
  NbJours := NbJEcart.value;
	FillChar(TabD,SizeOf(TabD),#0);
  Prevision := (EtatType = 'VENTIL') and (TRadioButton(GetControl('BALPREV')).Checked);

	// Calcul des périodes
	if Prevision then
		begin  			// si prévision, date d'arrete = date depart
    TabD[1] := StrToDate(DateArrete.Text) + 1;
    for i:=1 to NbPeriodes do
      begin
        TabD[i+4] := TabD[i] + NbJours;
        if i<4 then TabD[i+1] := TabD[i+4] + 1;
      end;
    end
 	else
  	begin				// sinon, date d'arrete = date de fin
    TabD[NbPeriodes+4] := StrToDate(DateArrete.Text);
    for i:=NbPeriodes downto 1 do
      begin
        TabD[i]:=TabD[i+4]-(NbJours-1);
        if i>1 then TabD[i+3]:=TabD[i]-1;
      end;
    end;

	// Remplissage des contrôles
  for i:=1 to 8 do
  	SetControlText('PERIODE'+IntToStr(i),DateToStr(TabD[i]));
end;

procedure TOF_CPTPAYEURGLAGE.CalculPeriodesAvecPeriodicite;
var
	i, ChoixPeriode : Integer ;
  an,mois,jour,JMax : Word ;
  DAT : TDATETIME ;
  Prevision : Boolean;
	TabD : TTabDate8;
begin
  // Initalisation et paramétrage suivant type état
  FillChar(TabD,SizeOf(TabD),#0);
  ChoixPeriode := Periodicite.itemindex;

  Prevision := (EtatType = 'VENTIL') and (TRadioButton(GetControl('BALPREV')).Checked);
	if Prevision
  	then TabD[1] := StrToDate(DateArrete.Text) + 1		 			// si prévision, date d'arrete = date depart
  	else TabD[NbPeriodes+4] := StrToDate(DateArrete.Text);	// sinon, date d'arrete = date de fin

	// Calcul en fonction de la périodicité choisie
  case ChoixPeriode of
    0,1,2,3,4,5 :			  			// ===> Périodicité de 1 à 6 mois
    	if Prevision then
        for i:=1 to NbPeriodes do  { Pour chaque Fourchette de dates( en partant de la première) , en mensuel }
          begin
          	TabD[i+4]:=PlusMois(TabD[i],(ChoixPeriode+1))-1;
            if i<NbPeriodes then TabD[i+1]:=TabD[i+4]+1 ;
          end
      else
        for i:=NbPeriodes downto 1 do { Pour chaque Fourchette de dates( en partant de la derniére) , en mensuel }
          begin
          DAT:=PlusMois(TabD[i+4],-(ChoixPeriode+1)) ;
          DecodeDate(DAT,an,mois,jour) ;
          JMax:=StrToInt(FormatDateTime('d',FinDeMois(EncodeDate(an,mois,1)))) ;
          TabD[i]:=PlusMois(TabD[i+4],-(ChoixPeriode+1))+(JMax-jour)+1 ;
          if i>1 then TabD[i+3]:=TabD[i]-1 ;
          end ;
    6           :					    // ===> Périodicité de 15 jours
    	if Prevision then
        for i:=1 to NbPeriodes do	{ Pour chaque Fourchette de dates (en partant de la première) , en Quinzaine }
           begin
           TabD[i+4]:=TabD[i]+ 15 ;
           if i<4 then TabD[i+1]:=TabD[i+4]+1 ;
           end
      else
        for i:=4 downto 1 Do { Pour chaque Fourchette de dates( en partant de la derniére) , en Quinzaine }
           begin
           DAT:=TabD[i+4] ;
           DecodeDate(DAT,an,mois,jour) ;
           JMax:=StrToInt(FormatDateTime('d',FinDeMois(EncodeDate(an,mois,1)))) ;
           if Jour<=15
           	then TabD[i]:=TabD[i+4]-15+1 			{ Date départ = (Date d'arrivée - 15 jours) + 1 jour, si date avant le 15 du mois }
           	else TabD[i]:=TabD[i+4]-(JMax-15)+1 ; { Date départ = (Date d'arrivée - (Nb jours Max du mois - 15 jours) ) + 1 jour, si date aprés le 15 du mois }
           if i>1 then TabD[i+3]:=TabD[i]-1 ;
           end;
    7           :             // ===> Périodicité de 1 semaine
    	if Prevision then
        for i:=1 to NbPeriodes Do	{ Pour chaque Fourchette de dates( en partant de la première) , en Hebdo }
          begin										{ Date d'arrivée = (Date de départ + 7 jours)  }
          TabD[i+4]:=TabD[i]+7;
          if i<4 then TabD[i+1]:=TabD[i+4]+1 ;
          end
			else
		    for i:=4 downto 1 Do  { Pour chaque Fourchette de dates( en partant de la derniére) , en Hebdo }
          begin            		{ Date départ = (Date d'arrivée - (7 jours+ 1 jours) )  }
          TabD[i]:=TabD[i+4]-(7)+1 ;
          if i>1 then TabD[i+3]:=TabD[i]-1 ;
          end ;
  	end ;
		// Remplissage des contrôles
    for i:=1 to 8 do	SetControlText('PERIODE'+IntToStr(i),DateToStr(TabD[i]));
end;

procedure TOF_CPTPAYEURGLAGE.BalTypeOnChange(Sender: TObject);
begin
	DateArreteOnChange(nil);
end;

procedure TOF_CPTPAYEURGLAGE.OnNew;
begin
  inherited;
(*FP
	if EtatType = 'VENTIL' then
  	begin
    // Changement de la nature de l'état
    TFQRS1(Ecran).NatureEtat := 'BAV'; // VL Erreur sur la nature de l'état
		TFQRS1(Ecran).CodeEtat := 'BAV';
    end ;*)
end;

procedure TOF_CPTPAYEURGLAGE.UpdateOptions(vBoAvecRupt: Boolean);
// MAJ accès aux checkbox liés aux options de ruptures
begin
  if not vBoAvecRupt then
    begin
    SetControlChecked( 'SAUTPAGE',    False );
    SetControlChecked( 'SURRUPTURE',  False );
    end ;
  SetControlChecked( 'CPTASSOCIES', vBoAvecRupt );
  SetControlEnabled( 'CPTASSOCIES', vBoAvecRupt );
  SetControlEnabled( 'SAUTPAGE',    vBoAvecRupt );
  SetControlEnabled( 'SURRUPTURE',  vBoAvecRupt );
end;

procedure TOF_CPTPAYEURGLAGE.OnClickChoixTaLi(Sender: TObject);
begin
  SetControlText('RUPTURE1','');
  SetControlText('RUPTURE2','');
  SetControlText('TRIPAR','');
  SetControlEnabled('TRIPAR',(ChoixTypeTaLi.Value <> 'SCL') and (ChoixTypeTaLi.Value <> 'CLI'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 12/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTPAYEURGLAGE.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;


Initialization
  registerclasses ( [ TOF_CPTPAYEURGLAGE ] ) ;
end.
