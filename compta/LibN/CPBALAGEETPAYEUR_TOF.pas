{***********UNITE*************************************************
Auteur  ...... : Frédéric BELLARD
Créé le ...... : 30/08/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPBALAGEETPAYEUR ()
Mots clefs ... : TOF;CPBALAGEETPAYEUR
*****************************************************************}
Unit CPBALAGEETPAYEUR_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     eMul,
     eQRS1,MaineAGL, UTob,
{$else}
     db,
     dbtables,
     mul,
     Fe_Main,     // AGLLanceFiche
     QRS1,
 {$ENDIF}

     forms,
     SysUtils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
   //fred
     Ent1,
     ParamSoc,		// GetParamSocSecur YMO
     spin;

(*TofMeth, HTB97, Filtre*)


procedure CPLanceFiche_BalAgeeTPayeur;
procedure CPLanceFiche_BalVentilTPayeur;

Type
  TOF_CPBALAGEETPAYEUR = Class (TOF)
  private
    NbPeriodes	     : Integer;
    TypEcr           : THMultiValComboBox;
    LibTypEcr        : THEdit;
    NatureTiers      : THValComboBox;
    Periodes         : TTabSheet;
    Periodicite      : THValComboBox ;
    DateArr	     : THEdit;
    Ecart            : THRadioGroup ;
    NbJEcart         : TSpinEdit ;
    CptCol1, CptCol2 : THEdit;
    MPaie            : THEdit;
    ModePaiement     : THMultiValComboBox ;
    Etablissement    : THValComboBox ;
    SorteSens        : THValComboBox ;
    DevPivot         : THEdit ;
    Devise           : THValComboBox ;
    AffichageMonnaie : THRadioGroup ;
    Resolution       : THRadioGroup ;
    AffMontants      : THValComboBox;
    ChoixTabLibres   : THRadioGroup;
    ConditionSQL     : THEdit;
    ConditionSQLCUMUL: THEdit;
    EtatType	     : String;
    BalPrev, BalRet  : TRadioButton;

    function  GetMinMaxCompte(stTable, stCol, stDefaut : String) : String;
    procedure DeviseOnChange (Sender : TObject) ;
    procedure ModePaiementOnChange (Sender : TObject);
    procedure EtablissementOnChange (Sender : TObject);
    procedure SorteSensOnChange (Sender : TObject);
    procedure PeriodiciteOnChange(Sender : TObject);
    procedure CalculPeriodesAvecPeriodicite;
    procedure DateArrOnChange(Sender : TObject);
    procedure EcartOnChange(Sender : TObject);
    procedure NbJEcartOnChange(Sender : TObject);
    procedure CalculPeriodesAvecNbJour;
    procedure NatureTiersOnChange (Sender : TObject) ;
    function GetTabletteColl(stNatureTiers : String): String;
    function GetTabletteTiers(stNatureTiers: String): String;
    procedure OnClickChoixTabLibres(Sender: TObject);
    procedure AffichageOnClick (Sender : TObject) ;
    procedure ImpressionOnClick (Sender : TObject);
    procedure AffMontantsOnChange (Sender : TObject);
    procedure CptCol1OnChange (Sender : TObject);
    procedure CptCol2OnChange (Sender : TObject);
    procedure Rupture1OnDblClick (Sender : TObject);
    procedure UpdateOptions(vBoAvecRupt: Boolean);
    procedure TriParOnChange (Sender : TObject);
    procedure TriParOnDblClick (Sender : TObject);
    procedure GenererConditionSQL;
    function GetWhereColl(stNatureTiers : String): String;
    procedure ParametreEtatType;
    procedure BalTypeOnChange(Sender : TObject);
    procedure AuxiElipsisClick(Sender : TObject);
  public
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (Arguments : string) ; override ;
  end ;

Type TTabDate8 = Array[1..8] of TDateTime ;

Implementation

uses TofTLTiersTri, TofTLTiersRupt, UTofMulParamGen; {26/04/07 YMO F5 sur Auxiliaire }

procedure CPLanceFiche_BalAgeeTPayeur;
begin
  // Etat non utilisable sous ORACLE7
  if V_PGI.Driver = dbORACLE7 then
    PGIInfo( 'Cet état n''est pas utilisable sous Oracle 7. Veuillez utiliser l''état correspondant dans le menu Autres éditions' , 'Balance Agée par Tiers payeur' )
  else
    AGLLanceFiche('CP', 'CPBALAGEETPAYEUR', '', '', 'AGEE') ;
end;

procedure CPLanceFiche_BalVentilTPayeur;
begin
  // Etat non utilisable sous ORACLE7
  if V_PGI.Driver = dbORACLE7 then
    PGIInfo( 'Cet état n''est pas utilisable sous Oracle 7. Veuillez utiliser l''état correspondant dans le menu Autres éditions' , 'Balance Ventilée par Tiers payeur' )
  else
    AGLLanceFiche('CP', 'CPBALAGEETPAYEUR', '', '', 'VENTIL') ;
end;

procedure TOF_CPBALAGEETPAYEUR.OnUpdate ;
Var
  St,St1,St2,St3,St4,St5,S,S1,TriLibelle,Order,stTL : string ;
  i : integer;
  stIndice : string;
Begin
  S:='';
  S1:='';
  TriLibelle:='T_AUXILIAIRE' ;
  St:=GetControlText('TRI') ;
  St1:=GetControlText('RUPTURE1') ;
  St2:=GetControlText('RUPTURE2') ;
  if GetControlText('TRILIBELLE')='X' then
    TriLibelle:='T_LIBELLE'
  else
    TriLibelle:='T_AUXILIAIRE' ;

  // Construction condition sur ruptures tables libres
  // Codes utilisés pour la gestion des ruptures sur tables libres :
  //    # = non gérée
  //    - = gérée mais non sélectionnée
  //    * = tous les enregistrements de la table (sans fourchette de valeurs)
  //  sinon code de valeur début dans St3, fin dans St4
  if St1<>'' then
    for i:=0 to 9 do begin
      St3:=ReadTokenSt(St1) ;
      St4:=ReadTokenSt(St2) ;
      if (St3<>'-') and (St3>'#') and (St3<>'') then begin// uniquement les tables sélectionnées
        if (St3<>'*') then begin// Si on a une fourchette de valeurs
          if ChoixTabLibres.Value = 'AUX' then
            stTL := '( T_TABLE'+IntToStr(i)+'>="'+St3+'" AND T_TABLE'+IntToStr(i)+'<="'+St4+'" )'
          else if ChoixTabLibres.Value = 'CLI' then
            stTL := '( YTC_TABLELIBRETIERS'+IntToHex(i+1,1)+'>="'+St3+'" AND YTC_TABLELIBRETIERS'+IntToHex(i+1,1)+'<="'+St4+'" )'
          else if ChoixTabLibres.Value = 'SCL' then
            stTL := '( YTC_RESSOURCE'+IntToHex(i+1,1)+'>="'+St3+'" AND YTC_RESSOURCE'+IntToHex(i+1,1)+'<="'+St4+'" )';
          if not TCheckBox(GetControl('CPTASSOCIES')).Checked then begin
            if ChoixTabLibres.Value = 'AUX' then
              stTL := '( ' + stTL +' OR ( T_TABLE'+IntToStr(i)+'="" ) ) '
            else if ChoixTabLibres.Value = 'CLI' then
              stTL := '( ' + stTL +' OR ( YTC_TABLELIBRETIERS'+IntToHex(i+1,1)+'="" ) ) '
            else if ChoixTabLibres.Value = 'SCL' then
              stTL := '( ' + stTL +' OR ( YTC_RESSOURCE'+IntToHex(i+1,1)+'="" ) ) ' ;
            end;
          S := S + ' AND ' +  stTL ;
         end
        else // Sans fourchette de valeur
          if TCheckBox(GetControl('CPTASSOCIES')).Checked then begin
            if ChoixTabLibres.Value = 'AUX' then
              S := S + ' AND (T_TABLE'+IntToStr(i)+'<>"") '
            else if ChoixTabLibres.Value = 'CLI' then
              S := S + ' AND (YTC_TABLELIBRETIERS'+IntToHex(i+1,1)+'<>"") '
            else if ChoixTabLibres.Value = 'SCL' then
              S := S + ' AND (YTC_RESSOURCE'+IntToHex(i+1,1)+'<>"") ' ;
            end;
        end;
      end;

  // conditions tiers payeur
  if (GetControlText('TIERSPAYEUR')='')  and (GetControlText('TIERSPAYEUR_')<>'') then
    S:=S+' AND T_PAYEUR<="'+GetControlText('TIERSPAYEUR_')+'"' ;
  if (GetControlText('TIERSPAYEUR')<>'') and (GetControlText('TIERSPAYEUR_')<>'') then
    S:=S+' AND T_PAYEUR>="'+GetControlText('TIERSPAYEUR')+'" AND T_PAYEUR<="'+GetControlText('TIERSPAYEUR_')+'"' ;

  // conditions auxiliaire
  if (GetControlText('AUXILIAIRE')='')  and (GetControlText('AUXILIAIRE_')<>'') then
    S:=S+' AND T_AUXILIAIRE<="'+GetControlText('AUXILIAIRE_')+'"' ;
  if (GetControlText('AUXILIAIRE')<>'') and (GetControlText('AUXILIAIRE_')<>'') then
    S:=S+' AND T_AUXILIAIRE>="'+GetControlText('AUXILIAIRE')+'" AND T_AUXILIAIRE<="'+GetControlText('AUXILIAIRE_')+'"' ;

  // conditions nature auxi
  if TComboBox(GetControl('NATURETIERS')).itemindex<>0 then
    S:=S+' AND T_NATUREAUXI="'+GetControlText('NATURETIERS')+'"' ;

  // Mise en place dans le champ caché
  SetControlText('WHERE',S) ;

  // MAJ pour gestion des ruptures dans l'état
  for i:=0 to 9 do begin
    SetControlText('T0'+IntToStr(i),'') ;
    SetControlText('ORDER'+IntToStr(i),'') ;
    end ;
  i := 0 ;
  while St<>'' do begin
    St5:=ReadTokenSt(St) ;
    if (St5<>'') then begin
        stIndice := copy(St5,3,1);
      if ChoixTabLibres.Value = 'AUX' then
        SetControlText('ORDER'+IntToStr(i),'T_TABLE'+stIndice)
      else if ChoixTabLibres.Value = 'CLI' then
        SetControlText('ORDER'+IntToStr(i),'YTC_TABLELIBRETIERS'+IntToHex(StrToInt(stIndice),1))
      else if ChoixTabLibres.Value = 'SCL' then
        SetControlText('ORDER'+IntToStr(i),'YTC_RESSOURCE'+IntToHex(StrToInt(stIndice),1));
      SetControlText('T0'+IntToStr(i),St5) ;
      if (St='') then begin
        if ChoixTabLibres.Value = 'AUX' then
          S1:=S1+'T_TABLE'+copy(St5,3,1)
        else if ChoixTabLibres.Value = 'CLI' then
          S1:=S1+'YTC_TABLELIBRETIERS'+IntToHex(StrToInt(stIndice),1)
        else if ChoixTabLibres.Value = 'SCL' then
          S1:=S1+'YTC_RESSOURCE'+IntToHex(StrToInt(stIndice),1)
       end
      else begin
        if ChoixTabLibres.Value = 'AUX' then
          S1:=S1+'T_TABLE'+copy(St5,3,1)+','
        else if ChoixTabLibres.Value = 'CLI' then
          S1:=S1+'YTC_TABLELIBRETIERS'+IntToHex(StrToInt(stIndice),1)+','
        else if ChoixTabLibres.Value = 'SCL' then
          S1:=S1+'YTC_RESSOURCE'+IntToHex(StrToInt(stIndice),1)+',';
        end;
      end ;
      inc(i) ;
    end ;

  if S1<>'' then
    Order:=S1+','+TriLibelle
  else
    Order:=TriLibelle ;

  { Mise à jour des requêtes pour récupérer les libellés des données des tables libres }
  for i := 0 to 9 do begin
    if ChoixTabLibres.Value = 'AUX' then
      SetControlText ('SQL'+IntToStr(i),'select nt_libelle FROM natcpte WHERE nt_typecpte="T0'+IntToStr(i)+'" and nt_nature=')
    else if ChoixTabLibres.Value = 'CLI' then
      SetControlText ('SQL'+IntToStr(i),'select yx_libelle FROM choixext WHERE yx_type="LT'+IntToHex(i+1,1)+'" and yx_code=')
    else if ChoixTabLibres.Value = 'SCL' then
      SetControlText ('SQL'+IntToStr(i),'select ars_libelle FROM ressource WHERE ars_ressource=');
    end;
  TFQRS1(Ecran).WhereSQL:='|'+Order ;

end ;

function TOF_CPBALAGEETPAYEUR.GetMinMaxCompte(stTable, stCol, stDefaut : String) : String;
var
  Q : TQuery;
  stWhere : String;
begin
  if stTable = 'GENERAUX' then
    if NatureTiers.Value <> '' then
      stWhere := GetWhereColl(NatureTiers.Value)
    else
      stWhere := ' WHERE G_COLLECTIF="X"'
  else
    if NatureTiers.Value <> '' then
      stWhere := ' WHERE T_NATUREAUXI="' + NatureTiers.Value + '"'
    else
      stWhere := ' WHERE T_PAYEUR<>''';
      Q := OpenSQL('SELECT ' + stCol +' CODE FROM ' + stTable + stWhere,True);
  if not Q.Eof then
    Result := Q.FindField('CODE').asString
  else
    Result := stDefaut;
  Ferme(Q);
end;

procedure TOF_CPBALAGEETPAYEUR.GenererConditionSQL;
var
  stWhere, stWhere2, stWhereQP, stTypeEcr, stMPaie : String;
begin
  // Condition sur les comptes tiers
  stWhere2 := ' AND E_AUXILIAIRE>="' + GetControlText('AUXILIAIRE') +
  					 '" AND E_AUXILIAIRE<="' + GetControlText('AUXILIAIRE_') + '"';

  //Condition sur les comptes tiers payeurs
  stWhere2 := stWhere2 + ' AND E_CONTREPARTIEAUX>="' + GetControlText('TIERSPAYEUR') +
  					 '" AND E_CONTREPARTIEAUX<="' + GetControlText('TIERSPAYEUR_') + '"';

  stWhere := stWhere + ' AND E_GENERAL>="' + GetControlText('COL1') +
  					 '" AND E_GENERAL<="' + GetControlText('COL2') + '"';

  if ModePaiement<>nil then
    stMPaie := ModePaiement.Value;
  if stMPaie <> '' then begin
    stWhereQP := '';
    while stMPaie <> '' do
      if stWhereQP = ''	then
        stWhereQP := ' AND E_MODEPAIE IN ("' + ReadTokenSt(stMPaie) + '"'
      else stWhereQP := stWhereQP + ',"' + ReadTokenSt(stMPaie) + '"';
    stWhere := stWhere + stWhereQP + ')';
    end;

  // Condition sur l'etablissement
  if Etablissement.itemIndex > 0 then
    stWhere := stWhere + ' AND E_ETABLISSEMENT = "' + Etablissement.Value + '"';

  // Condition sur la devise
  if Devise.itemIndex > 0 then
    stWhere := stWhere + ' AND E_DEVISE = "' + Devise.value + '"';

  // Condition sur le type de piece
  stWhere := stWhere + ' AND E_QUALIFPIECE<>"C" AND E_ECRANOUVEAU<>"CLO"'
                     + ' AND E_ECRANOUVEAU<>"OAN"' ;

  if TypEcr <> nil then
    stTypeEcr := TypEcr.Value;
  if stTypeEcr <> '' then begin
    stWhereQP := '';
    while stTypeEcr <> '' do
      if stWhereQP = ''	then
        stWhereQP := ' AND E_QUALIFPIECE IN ("' + ReadTokenSt(stTypeEcr) + '"'
      else stWhereQP := stWhereQP + ',"' + ReadTokenSt(stTypeEcr) + '"';
    stWhere := stWhere + stWhereQP + ')';
    end;

  //=======================
  ConditionSQL.Text := stWhere2 + stWhere;
  ConditionSQLCumul.Text := stWhere;

  end;

function TOF_CPBALAGEETPAYEUR.GetWhereColl(stNatureTiers : String): String;
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
    result := result + ' AND (G_NATUREGENE="COC" OR G_NATUREGENE="COD")' 		else
  if stNatureTiers='AUD' then
    result := result + ' AND (G_NATUREGENE="COF" OR G_NATUREGENE="COD")';
end;

procedure TOF_CPBALAGEETPAYEUR.OnLoad ;
var
  stTypeEcr : String;
begin
  inherited;

  if (GetControlText('TIERSPAYEUR')='') then
    SetControlText('TIERSPAYEUR',GetMinMaxCompte('TIERS', 'MIN(T_PAYEUR)','0'));
  if (GetControlText('TIERSPAYEUR_')='') then
    SetControlText('TIERSPAYEUR_',GetMinMaxCompte('TIERS', 'MAX(T_PAYEUR)','ZZZZZZZZZZZZZZZZZ'));

  if (GetControlText('AUXILIAIRE')='') then
    SetControlText('AUXILIAIRE',GetMinMaxCompte('TIERS', 'MIN(T_AUXILIAIRE)','0'));
  if (GetControlText('AUXILIAIRE_')='') then
    SetControlText('AUXILIAIRE_',GetMinMaxCompte('TIERS', 'MAX(T_AUXILIAIRE)','ZZZZZZZZZZZZZZZZZ'));

  if (CptCol1.Text='') then begin
    CptCol1.Text := GetMinMaxCompte('GENERAUX', 'MIN(G_GENERAL)','0');
    SetControlText('COL1',GetControlText('GENERAL'));
    end;
  if (CptCol2.Text='') then begin
    CptCol2.Text := GetMinMaxCompte('GENERAUX', 'MAX(G_GENERAL)','ZZZZZZZZZZZZZZZZZ');
    SetControlText('COL2',GetControlText('GENERAL_'));
    end;

  if AffMontants.ItemIndex = 1 then
    SetControlText('MONTANTS','SIGNES')
  else
    SetControlText('MONTANTS','NORMAL');


  LibTypEcr.Text := '';
  stTypeEcr := TypEcr.Value;
  if stTypeEcr <> '' then
    while stTypeEcr <> '' do
      LibTypEcr.Text := LibTypEcr.Text
      	+ RechDom(TypEcr.DataType,ReadTokenSt(stTypeEcr),False) + ' '
  else
    LibTypEcr.Text := '<<Tous>>';

  if (VH^.ExoV8.Code <> '') then
    SetControlText('DateExoV8', UsDateTime(VH^.ExoV8.Deb) )
  else
    SetControlText('DateExoV8', UsDateTime(iDate1900) ) ;

  // Remplir les conditions SQL
  GenererConditionSQL;
end ;

procedure TOF_CPBALAGEETPAYEUR.DeviseOnChange (Sender : TObject);
begin
  if Devise.ItemIndex=0 then begin
    SetControlText('DEVISEPIVOT','%');
    AffichageMonnaie.ItemIndex := 0;
    AffichageMonnaie.Enabled := False;
   end
  else begin
    SetControlText('DEVISEPIVOT',GetControlText('SORTEDEVISE'));
    AffichageMonnaie.Enabled := True;
    end;
end;

procedure TOF_CPBALAGEETPAYEUR.ModePaiementOnChange (Sender : TObject);
begin
  if ModePaiement.Value='<<Tous>>' then
    SetControlText('MPAIE','%')
  else
    SetControlText('MPAIE',GetControlText('E_MODEPAIE'));
end;

procedure TOF_CPBALAGEETPAYEUR.EtablissementOnChange (Sender : TObject);
begin
  if Etablissement.ItemIndex=0 then
    SetControlText('ETABL','%')
  else
    SetControlText('ETABL',GetControlText('ETABLISSEMENT'));
end;

procedure TOF_CPBALAGEETPAYEUR.SorteSensOnChange (Sender : TObject);
begin
  if SorteSens.ItemIndex=0 then
    SetControlText('SENS','%')
  else
    SetControlText('SENS',GetControlText('SORTESENS'));
end;

procedure TOF_CPBALAGEETPAYEUR.PeriodiciteOnChange(Sender : TObject);
begin
  inherited ;
  CalculPeriodesAvecPeriodicite;
end;

procedure TOF_CPBALAGEETPAYEUR.CalculPeriodesAvecPeriodicite;
var
  i, ChoixPeriode   : Integer ;
  an,mois,jour,JMax : Word ;
  DAT               : TDATETIME ;
  TabD              : TTabDate8;
  Prevision         : Boolean;
begin
  // Initialisation et paramétrage suivant type état
  FillChar(TabD,SizeOf(TabD),#0);
  ChoixPeriode := Periodicite.itemindex;
  Prevision := (EtatType = 'VENTIL') and (TRadioButton(GetControl('BALPREV')).Checked);
  if Prevision then
    TabD[1] := StrToDate(DateArr.Text) + 1
  else
    TabD[NbPeriodes+4] := StrToDate(DateArr.Text);	// sinon, date d'arrete = date de fin

  // Calcul en fonction de la périodicité choisie
  case ChoixPeriode of
    0,1,2,3,4 :			  			// ===> Périodicité de 1 à 6 mois
      if Prevision then begin
        for i:=1 to NbPeriodes do begin { Pour chaque Fourchette de dates( en partant de la première) , en mensuel }
          TabD[i+4]:=PlusMois(TabD[i],(ChoixPeriode+1))-1;
          if i<NbPeriodes then
            TabD[i+1]:=TabD[i+4]+1 ;
          end;
       end
      else begin
        for i:=NbPeriodes downto 1 do begin{ Pour chaque Fourchette de dates( en partant de la derniére) , en mensuel }
          DAT:=PlusMois(TabD[i+4],-(ChoixPeriode+1)) ;
          DecodeDate(DAT,an,mois,jour) ;
          JMax:=StrToInt(FormatDateTime('d',FinDeMois(EncodeDate(an,mois,1)))) ;
          TabD[i]:=PlusMois(TabD[i+4],-(ChoixPeriode+1))+(JMax-jour)+1 ;
          if i>1 then
            TabD[i+3]:=TabD[i]-1 ;
          end;
       end;
    6           :
      if Prevision then begin
        for i:=1 to NbPeriodes do begin	{ Pour chaque Fourchette de dates (en partant de la première) , en Quinzaine }
          TabD[i+4]:=TabD[i]+ 15 ;
          if i<4 then
            TabD[i+1]:=TabD[i+4]+1 ;
          end;
       end
      else begin				    // ===> Périodicité de 15 jours
        for i:=4 downto 1 do begin{ Pour chaque Fourchette de dates( en partant de la derniére) , en Quinzaine }
          DAT:=TabD[i+4] ;
          DecodeDate(DAT,an,mois,jour) ;
          JMax:=StrToInt(FormatDateTime('d',FinDeMois(EncodeDate(an,mois,1)))) ;
          if Jour<=15 then
            TabD[i]:=TabD[i+4]-15+1 { Date départ = (Date d'arrivée - 15 jours) + 1 jour, si date avant le 15 du mois }
          else
            TabD[i]:=TabD[i+4]-(JMax-15)+1 ; { Date départ = (Date d'arrivée - (Nb jours Max du mois - 15 jours) ) + 1 jour, si date aprés le 15 du mois }
          if i>1 then
            TabD[i+3]:=TabD[i]-1 ;
          end;
        end;
    7           :             // ===> Périodicité de 1 semaine
    	if Prevision then begin
          for i:=1 to NbPeriodes do begin  { Pour chaque Fourchette de dates( en partant de la première) , en Hebdo }
            TabD[i+4]:=TabD[i]+7;	   { Date d'arrivée = (Date de départ + 7 jours)  }
            if i<4 then
              TabD[i+1]:=TabD[i+4]+1 ;
            end;
         end
        else begin
          for i:=4 downto 1 do begin  { Pour chaque Fourchette de dates( en partant de la derniére) , en Hebdo }
            TabD[i]:=TabD[i+4]-(7)+1 ;  { Date départ = (Date d'arrivée - (7 jours+ 1 jours) )  }
            if i>1 then
              TabD[i+3]:=TabD[i]-1 ;
            end;
          end;
  end;

  // Remplissage des contrôles
  for i:=1 to 8 do
    SetControlText('PERIODE'+IntToStr(i),DateToStr(TabD[i]));
end;

procedure TOF_CPBALAGEETPAYEUR.EcartOnChange(Sender : TObject);
begin
  inherited ;
  SetControlVisible('PERIODES',Ecart.ItemIndex=1);
  if Ecart.ItemIndex=0 then begin
    Periodes.TabVisible:=false;
    NbJEcart.Enabled:=true;
    NbJEcartOnChange(nil)
   end
  else begin
    NbJEcart.Enabled:=false;
    PeriodiciteOnChange(nil);
    end;
end;

procedure TOF_CPBALAGEETPAYEUR.NbJEcartOnChange(Sender : TObject);
begin
  inherited;
  if GetControlText('NBJECART')='' then
    exit;
  CalculPeriodesAvecNbJour;
end;

procedure TOF_CPBALAGEETPAYEUR.CalculPeriodesAvecNbJour;
var
  i, NbJours : Integer ;
  TabD       : TTabDate8;
  Prevision  : boolean;
begin
  NbJours := NbJEcart.value;
  FillChar(TabD,SizeOf(TabD),#0);

  Prevision := (EtatType = 'VENTIL') and (TRadioButton(GetControl('BALPREV')).Checked);

  // Calcul des périodes
  if Prevision then begin // si prévision, date d'arrete = date depart
    TabD[1] := StrToDate(DateArr.Text) + 1;
    for i:=1 to NbPeriodes do begin
      TabD[i+4] := TabD[i] + NbJours;
      if i<4 then
        TabD[i+1] := TabD[i+4] + 1;
     end;
   end
  else begin				// sinon, date d'arrete = date de fin
    TabD[NbPeriodes+4] := StrToDate(DateArr.Text);
    for i:=NbPeriodes downto 1 do begin
      TabD[i]:=TabD[i+4]-(NbJours-1);
      if i>1 then
        TabD[i+3]:=TabD[i]-1;
     end;
   end;

  // Remplissage des contrôles
  for i:=1 to 8 do
    SetControlText('PERIODE'+IntToStr(i),DateToStr(TabD[i]));
end;

procedure TOF_CPBALAGEETPAYEUR.DateArrOnChange(Sender : TObject);
begin
  inherited ;
  if not IsValidDate(GetControlText('DATEARR')) then
    exit;
  if (EtatType = 'VENTIL') and (TRadioButton(GetControl('BALPREV')).Checked) then
    SetControlText('PERIODE1',GetControlText('DATEARR'))
  else
    SetControlText('PERIODE'+IntToStr(NbPeriodes+4),GetControlText('DATEARR'));
  EcartOnChange(nil);
end;

procedure TOF_CPBALAGEETPAYEUR.NatureTiersOnChange (Sender : TObject) ;
var
  CGen, CAux : string ;
begin
  inherited ;
  SetControlText('TIERSPAYEUR','');
  SetControlText('TIERSPAYEUR_','');
  SetControlText('AUXILIAIRE','');
  SetControlText('AUXILIAIRE_','');
  SetControlText('GENERAL','');
  SetControlText('GENERAL_','');
  SetControlText('NATUREAUX',GetControlText('NATURETIERS'));
  CGen := GetTabletteColl(NatureTiers.value);
  CAux := GetTabletteTiers(NatureTiers.value);
  SetControlProperty('TIERSPAYEUR','DATATYPE',CAux);
  SetControlProperty('TIERSPAYEUR_','DATATYPE',CAux);
  SetControlProperty('AUXILIAIRE','DATATYPE',CAux);
  SetControlProperty('AUXILIAIRE_','DATATYPE',CAux);
  SetControlProperty('GENERAL','DATATYPE',CGen);
  SetControlProperty('GENERAL_','DATATYPE',CGen);
end;

function TOF_CPBALAGEETPAYEUR.GetTabletteColl(stNatureTiers : String): String;
begin
  if stNatureTiers='CLI' then result := 'TZGCOLLCLIENT'
  else if stNatureTiers='FOU' then result := 'TZGCOLLFOURN'
  else if stNatureTiers='SAL' then result := 'TZGCOLLSALARIE'
  else if stNatureTiers='DIV' then result := 'TZGCOLLDIVERS'
  else if stNatureTiers='AUC' then result := 'TZGCOLLTOUTDEBIT'
  else if stNatureTiers='AUD' then result := 'TZGCOLLTOUTCREDIT'
  else result := 'TZGCOLLECTIF';
end;

function TOF_CPBALAGEETPAYEUR.GetTabletteTiers(stNatureTiers: String): String;
begin
  if stNatureTiers='CLI' then result := 'TZTCLIENT'
  else if stNatureTiers='FOU' then result := 'TZTFOURN'
  else if stNatureTiers='SAL' then result := 'TZTSALARIE'
  else if stNatureTiers='DIV' then result := 'TZTDIVERS'
  else if stNatureTiers='AUC' then result := 'TZTTOUTDEBIT'
  else if stNatureTiers='AUD' then result := 'TZTTOUTCREDIT'
  else result := 'TZTTOUS';
end;

procedure TOF_CPBALAGEETPAYEUR.OnClickChoixTabLibres(Sender: TObject);
begin
  SetControlText('RUPTURE1','');
  SetControlText('RUPTURE2','');
  SetControlText('TRI','');
  SetControlEnabled('TRI',(ChoixTabLibres.Value <> 'SCL') and (ChoixTabLibres.Value <> 'CLI'));
end;

procedure TOF_CPBALAGEETPAYEUR.AffichageOnClick (Sender : TObject) ;
begin
case AffichageMonnaie.ItemIndex of
  0 : begin
        SetControlText('DEBIT','E_DEBIT');
        SetControlText('CREDIT','E_CREDIT');
        SetControlText('COUVERTURE','E_COUVERTURE');
      end;
  1 : begin
        SetControlText('DEBIT','E_DEBITDEV');
        SetControlText('CREDIT','E_CREDITDEV');
        SetControlText('COUVERTURE','E_COUVERTUREDEV');
      end;
  end;
end;

procedure TOF_CPBALAGEETPAYEUR.ImpressionOnClick (Sender : TObject);
begin
case Resolution.ItemIndex of
  0 : SetControlText('FORMAT','DECI');
  1 : SetControlText('FORMAT','KILO');
  2 : SetControlText('FORMAT','MEGA');
  3 : SetControlText('FORMAT','ENTI');
  end;
end;

procedure TOF_CPBALAGEETPAYEUR.AffMontantsOnChange (Sender : TObject);
begin
  if AffMontants.ItemIndex=0 then
    SetControlText('MONTANTS','%')
  else
    SetControlText('MONTANTS',GetControlText('SORTEMONTANTS'));
end;

procedure TOF_CPBALAGEETPAYEUR.CptCol1OnChange (Sender : TObject);
begin
  SetControlText('COL1',GetControlText('GENERAL'));
end;

procedure TOF_CPBALAGEETPAYEUR.CptCol2OnChange (Sender : TObject);
begin
  SetControlText('COL2',GetControlText('GENERAL_'));
end;

procedure TOF_CPBALAGEETPAYEUR.Rupture1OnDblClick (Sender : TObject);
var
  St,St1,St2,St3,Arg : String ;
  i       : integer ;
  OkAssoc : integer ;
begin
  OkAssoc := 0;
  Arg     := GetControlText('RUPTURE1')+'|'+GetControlText('RUPTURE2');
  if GetControlVisible('CHOIXTABLIBRES') then
    Arg := Arg + '|' + GetControlText('CHOIXTABLIBRES')
  else
    Arg:= Arg + '|AUX';
  St      := CPLanceFiche_TLTIERSRUPT(Arg);
  St1     := ReadTokenPipe(St,'|');
  SetControlText('RUPTURE1',St1);
  SetControlText('RUPTURE2',St);
  for i:=0 to 9 do begin
    St2:=ReadTokenSt(St1);
    if (St2<>'#') and (St2<>'-') and (St2<>'') then begin
      if ChoixTabLibres.Value = 'AUX' then
        St3:=St3+'T0'+IntToStr(i)+';'
      else if ChoixTabLibres.Value = 'CLI' then
        St3:=St3+'LT'+IntToHex(i+1,1)+';'
      else
        St3 :=St3+'AR'+IntToHex(i+1,1)+';';
      OkAssoc:=OkAssoc+1 ;
      end;
    SetControlText('TRI',St3);
    end;

  // MAJ interface si aucune rupture sélectionnée
  UpdateOptions( OkAssoc > 0 ) ;
  if OkAssoc = 0 then
    begin
    SetControlText('RUPTURE1', '');
    SetControlText('RUPTURE2', '');
    SetControlText('TRI',   '');
    end ;
end;

procedure TOF_CPBALAGEETPAYEUR.UpdateOptions(vBoAvecRupt: Boolean);
// MAJ accès aux checkbox liés aux options de ruptures
begin
  if not vBoAvecRupt then begin
    SetControlChecked( 'SAUTPAGE',    False );
    SetControlChecked( 'SURRUPTURE',  False );
    end ;
  SetControlChecked( 'CPTASSOCIES', vBoAvecRupt );
  SetControlEnabled( 'CPTASSOCIES', vBoAvecRupt );
  SetControlEnabled( 'SAUTPAGE',    vBoAvecRupt );
  SetControlEnabled( 'SURRUPTURE',  vBoAvecRupt );
end;

procedure TOF_CPBALAGEETPAYEUR.TriParOnChange (Sender : TObject);
begin
  if GetControlText('TRI')='' then begin
    SetControlChecked('SAUTPAGE',False);
    SetControlChecked('SURRUPTURE',False);
   end
  else begin
    SetControlEnabled('SAUTPAGE',True);
    SetControlEnabled('SURRUPTURE',True);
    end;
end;

procedure TOF_CPBALAGEETPAYEUR.TriParOnDblClick (Sender : TObject);
var
  St,Arg:String;
begin
  Arg := GetControlText('TRI');
  St  := CPLanceFiche_TLTIERSTRI(Arg);
  SetControlText('TRI',St);
end;

procedure TOF_CPBALAGEETPAYEUR.OnArgument (Arguments : string) ;
var
  Rupture1, Rupture2, Tri : THEdit;
begin
  Inherited ;
  // Mémorisation du type d'état (AGEE ou VENTIL)
  EtatType := Arguments;
  NbPeriodes := 4 ;

  TypEcr       := THMultiValComboBox(Getcontrol('TYPECRITURE'));
  LibTypEcr    := THEdit(GetControl('LIBTYPECR'));
  NatureTiers  := THValComboBox(GetControl('NATURETIERS'));
  Periodes     := TTabSheet(GetControl('PERIODES'));
  Periodicite  := THValComboBox(GetControl('PERIODICITE'));
  DateArr      := THEdit(GetControl('DATEARR'));
  Ecart        := THRadioGroup(GetControl('ECART'));
  NbJEcart     := TSpinEdit(GetControl('NBJECART'));
  BalPrev      :=TRadioButton(GetControl('BALPREV'));
  BalRet       :=TRadioButton(GetControl('BALRET'));

  CptCol1       := THEdit(GetControl('GENERAL'));
  CptCol2       := THEdit(GetControl('GENERAL_'));
  ModePaiement  := THMultiValComboBox(GetControl('E_MODEPAIE'));
  MPaie         := THEdit(GetControl('MPAIE'));
  Etablissement := THValComboBox(GetControl('ETABLISSEMENT'));
  SorteSens     := THValComboBox(GetControl('SORTESENS'));
  DevPivot      := THEdit(GetControlText('DEVISEPIVOT'));
  Devise        := THValComboBox(GetControl('SORTEDEVISE'));

  AffichageMonnaie := THRadioGroup(GetControl('AFFICHAGEMONNAIE')) ;
  Resolution    := THRadioGroup(GetControl('TRESOLUTION')) ;
  AffMontants   := THValComboBox(GetControl('SORTEMONTANTS')) ;

  Rupture1      := THEdit(GetControl('RUPTURE1'));
  Rupture2      := THEdit(GetControl('RUPTURE2'));
  Tri           := THEdit(GetControl('TRI'));
  ChoixTabLibres :=THRadioGroup(GetControl('CHOIXTABLIBRES'));

  ConditionSQL  := THEdit(GetControl('CONDITIONSQL'));
  ConditionSQLCumul  := THEdit(GetControl('CONDITIONSQLCUMUL'));

  if Devise<>nil then begin
    Devise.OnChange:=DeviseOnChange;
    Devise.itemindex:=0;
    Devise.OnChange(NIL);
    end;
  if ModePaiement<>nil then begin
    ModePaiement.OnChange:=ModePaiementOnChange;
    ModePaiement.Value:='<<Tous>>';
    ModePaiement.OnChange(NIL);
    end;
  if Etablissement<>nil then begin
    Etablissement.OnChange:=EtablissementOnChange;
    Etablissement.itemindex:=0;
    Etablissement.OnChange(NIL);
    end;
  if SorteSens<>nil then begin
    SorteSens.OnChange:=SorteSensOnChange;
    SorteSens.itemindex:=2;
    SorteSens.OnChange(NIL);
    end;
  if BalPrev<>nil then begin
    BalPrev.OnClick := BalTypeOnChange;
    BalPrev.Checked:=true;
    end;
  if BalRet<>nil then
    BalRet.OnClick  := BalTypeOnChange;

  Periodes.TabVisible := (Ecart.ItemIndex=1) ;
  if EtatType='VENTIL' then begin
    if BalPrev.Checked then
      Periodes.Caption:='Prévisions'
    else
      Periodes.Caption:='Retards';
   end
  else
    Periodes.Caption:='Périodes';

  if Periodicite<>nil then Begin
    Periodicite.OnChange:=PeriodiciteOnChange;
    Periodicite.itemindex:=0;
    end;
  if DateArr<>nil then begin
    DateArr.OnChange:=DateArrOnChange;
    SetControlText('DATEARR',DateToStr(V_PGI.DateEntree));
    end;
  if Ecart<>nil then begin
    Ecart.OnClick:=EcartOnChange;
    Ecart.itemindex:=1;
    EcartOnChange(NIL);
    end;
  if NbJEcart<>nil then begin
    NbJEcart.OnChange:=NbJEcartOnChange;
    NbJEcart.text:='30';
    end;;
  if NatureTiers<>nil then begin
    NatureTiers.OnChange:=NatureTiersOnChange;
    NatureTiers.itemindex:=0;
    NatureTiersOnChange(nil);
    end;
  if ChoixTabLibres <> nil then begin
    ChoixTabLibres.OnClick := OnClickChoixTabLibres;
    ChoixTabLibres.ItemIndex := 0;
    end;
  PositionneEtabUser(Etablissement, False);
  if TypEcr<>nil then begin
    TypEcr.Value := 'N';
    TypEcr.Enabled := false;
    end;
  if AffichageMonnaie<>nil then begin
    AffichageMonnaie.OnClick:=AffichageOnClick;
    AffichageMonnaie.itemindex:=0;
    SetControlText('DEBIT','E_DEBIT') ;
    SetControlText('CREDIT','E_CREDIT') ;
    SetControlText('COUVERTURE','E_COUVERTURE') ;
    end;
  if AffMontants<>nil then begin
    AffMontants.OnChange:=AffMontantsOnChange;
    AffMontants.itemindex:=0;
    AffMontants.OnChange(NIL);
    end;
  if Resolution<>nil then begin
    Resolution.OnClick:=ImpressionOnClick;
    Resolution.itemindex:=1;
    SetControlText('FORMAT','DECI');
    end;
  if (CptCol1<>nil) and (CptCol2<>nil) then begin
    CptCol1.OnChange:=CptCol1OnChange;
    CptCol2.OnChange:=CptCol2OnChange;
    end;
  if Rupture1<>nil then
    Rupture1.OnDblClick:=Rupture1OnDblClick;
  if Rupture2<>nil then
    Rupture2.OnDblClick:=Rupture1OnDblClick;
  if Tri<>nil then begin
    Tri.OnChange:=TriParOnChange;
    Tri.OnDblClick:=TriParOnDblClick;
    SetControlEnabled('SAUTPAGE',False);
    SetControlEnabled('SURRUPTURE',False);
    end;

  //=====> Param selon type de balance
  ParametreEtatType;

  // Init options des ruptures
  UpdateOptions(false);
  ChoixTabLibres.Visible := TL_TIERSCOMPL_Actif;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    THEdit(GetControl('AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;
    THEdit(GetControl('AUXILIAIRE_', true)).OnElipsisClick:=AuxiElipsisClick;
  end;

end ;

procedure TOF_CPBALAGEETPAYEUR.ParametreEtatType;
var
  stTitre : String;
begin
  if EtatType = 'VENTIL' then begin
    // Modif titre QRS1
    stTitre := 'Balance Ventilée par Tiers payeur';
    TFQRS1(Ecran).Caption := stTitre;
    UpdateCaption(Ecran);
    //Affiche du type de balance (Prevision / Retard)
    SetControlVisible('TBALTYPE',True);
    SetControlVisible('BALPREV',True);
    SetControlVisible('BALRET',True);
    SetControlProperty('BALRET','checked',False);
    SetControlProperty('BALPREV','checked',True);
    // Changement de la nature de l'état
    TFQRS1(Ecran).NatureEtat := 'TFA';
    TFQRS1(Ecran).CodeEtat := 'BVT';
   end
  else begin
    // Modif titre QRS1
    stTitre := 'Balance Agée par Tiers payeur';
    TFQRS1(Ecran).Caption := stTitre;
    UpdateCaption(Ecran);
    //Cache le type de balance ventilée (Prevision / Retard)
    SetControlVisible('TBALTYPE',False);
    SetControlVisible('BALPREV',False);
    SetControlVisible('BALRET',False);
    // Changement de la nature de l'état
    TFQRS1(Ecran).NatureEtat := 'TFA';
    TFQRS1(Ecran).CodeEtat := 'BAT';
    end;
  Ecran.Caption :=TraduireMemoire(Ecran.Caption) ;
  UpdateCaption(Ecran);
end;

procedure TOF_CPBALAGEETPAYEUR.BalTypeOnChange(Sender: TObject);
begin
  if BalPrev.Checked then
    Periodes.Caption:='Prévisions'
  else
    Periodes.Caption:='Retards';

  DateArrOnChange(nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 26/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPBALAGEETPAYEUR.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;


Initialization
  registerclasses ( [ TOF_CPBALAGEETPAYEUR ] ) ;
end.

