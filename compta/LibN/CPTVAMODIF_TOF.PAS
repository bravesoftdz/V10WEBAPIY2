{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/09/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPTVAMODIF ()
Mots clefs ... : TOF;CPTVAMODIF
*****************************************************************}
Unit CPTVAMODIF_TOF ;

Interface

Uses Hqry,
     StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     Hdb,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,         // AGLLanceFiche
{$ELSE}
     eMul,
     MainEAgl,        // AGLLanceFFiche
{$ENDIF}
{$IFDEF COMPTA}
     Saisie,		    // Pour Saisie eAGL
{$ENDIF}
     Utob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     Ent1,		              // Pour EstMonaieIn et GetPeriode
     Ed_Tools,              // Pour le videListe
     TiersPayeur,           // Pour les fonctions xxxTP
     HStatus,               // Pour la barre d'état
     Zcompte,
     SaisComm,
     dialogs,              // Pour les procédures de MAJ des comptes
     HSysMenu,
     ULibExercice,
     SaisBase,
     CPELEMENTSTVA_TOF,
     saisutil,
     AGLInit,
     utilPGI,   // _Blocage
     ParamDBG,  // ParamList
     Htb97,     // ToolBarButton97
     paramsoc;

Type
  TOF_CPTVAMODIF = Class (TOF)
   // Eléments interface
    JOURNAL        : THValComboBox ;
    NATUREPIECE    : THValComboBox ;
    DEVISE         : THValComboBox ;
    EXERCICE       : THValComboBox ;
    ETABLISSEMENT  : THValComboBox ;
    GENERAL        : THEdit ;
    DATECOMPTABLE  : THEdit ;
    DATECOMPTABLE_ : THEdit ;
    DATECREATION   : THEdit ;
    DATECREATION_  : THEdit ;
    DATEECHEANCE   : THEdit ;
    DATEECHEANCE_  : THEdit ;
    DATEPAQUETMAX  : THEdit ;
    XX_WHERE       : THEdit ;
    XX_WHEREINIT   : THEdit ;
    XX_SELECT      : THEdit ;
    XX_FROM        : THEdit ;
    Maj            : THEdit ;
    PAGES          : TPageControl;
    BParamListe    : TToolBarButton97;
    MsgBox    : THMsgBox;
  {$IFDEF EAGLCLIENT}
    FListe : THGrid ;
  {$ELSE}
    FListe : THDBGrid ;
  {$ENDIF}
    FEcran    : TFMul ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnClickBParamListe   (Sender : TObject);
    procedure InitComposants (S : String) ;
    // Evènements fiche
    procedure EXERCICEChange(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    function  RecupValControl : String;
    function  ConstitueLigneOD(TobDetail : Tob) : String;
    function  ConstituePieceOD(TypeEcr : integer ; TobDetail : Tob) : String;
    function  ConstituePieceOrig(TypeEcr : integer ; Q : TQuery) : String;
    function  SensInverse(TobDet : Tob) : Boolean;
    procedure OuvrirClick ;
    procedure CopieEche(TypeEcr : integer ; TobDet : Tob);
    procedure OutilMaj;
    procedure ValideClick ;
    procedure InitMsgBox;
    procedure BValideClick ( Sender : TObject ) ;
    function  AddFille (T : TOB ; Q : TQuery) : TOB;
    Function  EstCeClient ( Gene,Auxi : String ) : boolean ;
    procedure RemplirOM ( OM : TMOD ; Q1 : TQuery ) ;
    procedure MajModTvaEnc ;
    procedure RefaitXXWHEREINIT ;
  private
    UpSQL : String ;
    NowFutur : TDateTime ;
    FStListeParam   : string;
    FNewSources     : string;
    FNewLiaison     : string;
    FStNewTris      : string;
    FNewParams      : string;
    FStNewTitres    : hstring;
    FNewJustifs     : string;
    FNewLibelle     : hstring;
    FNewNumCols     : hstring;
    FNewPerso       : string;

    FNewOkTri       : Boolean;
    FNewOkNumCol    : Boolean;

    FStListeChamps  : string;            // Liste des champs de la LISTE PARAMETRABLE
    FStNewLargeurs  : string;            // Largeur des colonnes dans la liste paramètrable
  end ;

function CPLanceFiche_ModifTVA( vStParam : string = '' ) : string ;
function CPLanceFiche_MajTVA( vStParam : string = '' ) : string ;

procedure TvaModifEnc;

Implementation

uses
//  RepDevEur, // ChangeLeTauxDevise
  UlibEcriture,
  {$IFDEF eAGLCLIENT}
  MenuOLX
  {$ELSE}
  MenuOLG
  {$ENDIF eAGLCLIENT}
  , Constantes
  , UTofMulParamGen; {26/04/07 YMO F5 sur Auxiliaire }

function CPLanceFiche_ModifTVA( vStParam : string = '' ) : string ;
begin
  Result := AGLLanceFiche('CP', 'CPTVAMODIF', '', '', vStParam);
end;

function CPLanceFiche_MajTVA( vStParam : string = '' ) : string ;
begin
  Result := AGLLanceFiche('CP', 'CPTVAMODIF', '', '', 'Maj');
end;

procedure TvaModifEnc;
begin
  if PasCreerDate(V_PGI.DateEntree) then Exit ;
  if _Blocage(['nrCloture','nrBatch','nrLettrage','nrEnca','nrDeca'],True,'nrSaisieModif') then Exit ;
  AGLLanceFiche('CP','CPTVAMODIF','','','');
end;

procedure TOF_CPTVAMODIF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPTVAMODIF.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPTVAMODIF.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPTVAMODIF.RefaitXXWHEREINIT ;
BEGIN
  XX_WHEREINIT.Text:=' (ECRITURE.E_GENERAL LIKE "'+GetParamSocSecur('SO_COLLCLIENC','')
                    +'%" OR ECRITURE.E_GENERAL LIKE "'+GetParamSocSecur('SO_COLLFOUENC','')+'%")';

  If Maj.Text ='X' then
    XX_WHEREINIT.Text:=XX_WHEREINIT.Text
      +' AND J_NATUREJAL="OD"'
      +' AND (ECRITURE.E_MODESAISIE = "" OR ECRITURE.E_MODESAISIE = "-")'
      +' AND (ECRITURE.E_ECHEENC1=0 AND ECRITURE.E_ECHEENC2=0'
      +' AND ECRITURE.E_ECHEENC3=0 AND ECRITURE.E_ECHEENC4=0 AND ECRITURE.E_ECHEDEBIT=0)'
      +' AND (E2.E_LETTRAGE IS NOT NULL AND E2.E_LETTRAGE<>"")'
      +' AND (E2.E_GENERAL LIKE "'+GetParamSocSecur('SO_COLLCLIENC','')+'%"'
      +' OR E2.E_GENERAL LIKE "'+GetParamSocSecur('SO_COLLFOUENC','')+'%")'
      + ' AND ( (ECRITURE.E_ETATLETTRAGE<>"TL" AND ECRITURE.E_ETATLETTRAGE<>"RI")'
      + ' OR (ECRITURE.E_ETATLETTRAGE="TL" AND ECRITURE.E_DATEPAQUETMAX>"'+ USDateTime(StrToDate(DATEPAQUETMAX.Text)) + '"))';
END ;

procedure TOF_CPTVAMODIF.OnLoad ;
begin
// Modif GP Bug sur date paquet max. (état toujours positionné dans le xxwhereinit à 01/01/2099 !!
  RefaitXXWHEREINIT ;
  XX_WHERE.Text:=XX_WHEREINIT.Text+RecupValControl;

  FStListeParam:=TFMul(Ecran).DbListe;

  ChargeHListe (FStListeParam, FNewSources, FNewLiaison, FStNewTris, FStListeChamps, FStNewTitres, FStNewLargeurs, FNewJustifs,
               FNewParams, FNewLibelle, FNewNumCols, FNewPerso, FNewOkTri, FNewOkNumCol);

  XX_SELECT.Text:=StringReplace(FStListeChamps,';', ', ',[rfReplaceAll]);
  XX_SELECT.Text:=StringReplace(XX_SELECT.Text,'E_', 'ECRITURE.E_',[rfReplaceAll]);
//  XX_SELECT.Text:=XX_SELECT.Text+'E2.E_GENERAL';
  XX_SELECT.Text:=copy(XX_SELECT.Text,1,Length(XX_SELECT.Text)-2);

  XX_FROM.Text:='ECRITURE';
  If Maj.Text='X' then XX_FROM.Text:=XX_FROM.Text
   +' LEFT OUTER JOIN JOURNAL ON ECRITURE.E_JOURNAL=J_JOURNAL'
   +' LEFT OUTER JOIN ECRITURE E2 ON E2.E_NUMEROPIECE=ECRITURE.E_NUMEROPIECE'
   +' AND E2.E_NUMLIGNE<>ECRITURE.E_NUMLIGNE AND E2.E_JOURNAL=ECRITURE.E_JOURNAL'
   +' AND E2.E_EXERCICE=ECRITURE.E_EXERCICE AND E2.E_DATECOMPTABLE=ECRITURE.E_DATECOMPTABLE'
   +' AND E2.E_QUALIFPIECE=ECRITURE.E_QUALIFPIECE AND E2.E_NUMECHE=ECRITURE.E_NUMECHE';

  Inherited ;
end ;

procedure TOF_CPTVAMODIF.OnArgument (S : String ) ;
begin
  Inherited ;

  JOURNAL        := THValComboBox(GetControl('ZZJOURNAL', True)) ;
  NATUREPIECE    := THValComboBox(GetControl('ZZNATUREPIECE', True)) ;
  DEVISE         := THValComboBox(GetControl('ZZDEVISE', True)) ;
  EXERCICE       := THValComboBox(GetControl('ZZEXERCICE', True)) ;
  ETABLISSEMENT  := THValComboBox(GetControl('ZZETABLISSEMENT', True)) ;
  GENERAL        := THEdit(GetControl('ZZGENERAL', True))  ;
  DATECOMPTABLE  := THEdit(GetControl('ZZDATECOMPTABLE', True))  ;
  DATECOMPTABLE_ := THEdit(GetControl('ZZDATECOMPTABLE_', True)) ;
  DATECREATION   := THEdit(GetControl('ZZDATECREATION', True))  ;
  DATECREATION_  := THEdit(GetControl('ZZDATECREATION_', True)) ;
  DATEECHEANCE   := THEdit(GetControl('ZZDATEECHEANCE', True)) ;
  DATEECHEANCE_  := THEdit(GetControl('ZZDATEECHEANCE_', True)) ;
  DATEPAQUETMAX  := THEdit(GetControl('ZZDATEPAQUETMAX', True)) ;
  XX_WHERE       := THEdit(GetControl('XX_WHERE', True)) ;
  XX_WHEREINIT   := THEdit(GetControl('XXWHEREINIT', True)) ;
  XX_SELECT      := THEdit(GetControl('XX_SELECT', True)) ;
  XX_FROM        := THEdit(GetControl('XX_FROM', True)) ;
  MAJ            := THEdit(GetControl('MAJ', True)) ;
  PAGES          := TPageControl(GetControl('Pages', True)) ;

  If S='Maj' then
  begin
    Ecran.caption:='Mise à jour des éléments de TVA';
    UpdateCaption(TFMul(Ecran));
    Maj.Text:='X';
    TFMul(Ecran).SetDBListe('CPTVAMODIF');
  end
  else
    TFMul(Ecran).SetDBListe('TVAMODIF');

  FEcran := TFMul(Ecran) ;

{$IFDEF EAGLCLIENT}
  FListe  := THGrid( GetControl('FListe',True) ) ;
{$ELSE}
  FListe  := THDBGrid( GetControl('FListe',True)) ; //THDB YMOO
{$ENDIF}

  EXERCICE.OnChange  := EXERCICEChange ;

  FListe.OnDblClick 	:= FListeDblClick ;
//ymoo  FListe.ListeParam   := FStListeParam;
//  BParamListe.OnClick := OnClickBParamListe;

  TButton(GetControl('BValider', True)).OnClick := BValideClick ;
  TButton(GetControl('BZOOMPIECE',True)).OnClick := BZoomPieceClick;

  InitComposants(S);
end ;

procedure TOF_CPTVAMODIF.OnClickBParamListe(Sender: TObject);
begin
  if V_Pgi.SAV then
  begin
    PgiInfo('Vous ne pouvez pas modifier la liste des champs en mode SAV.', Ecran.Caption);
    Exit;
  end;

{$IFDEF EAGLCLIENT}
  ParamListe(FStListeParam, nil, 'Personnalisation des listes');
{$ELSE}
  ParamListe(FStListeParam, nil, nil, 'Personnalisation des listes');
{$ENDIF}


  //TraitementListeParam;

//  FMulPanelFiltre.InitializeField;
//  FMulPanelFiltre.InitializeControl;

//  BCherche.Click;
end;

procedure TOF_CPTVAMODIF.InitComposants (S : String);
begin
  if VH^.Precedent.Code<>'' then
      DATECOMPTABLE.Text := DateToStr(VH^.Precedent.Deb)
  else
      DATECOMPTABLE.Text := DateToStr(VH^.Encours.Deb) ;

  DATECOMPTABLE_.Text := DateToStr(V_PGI.DateEntree) ;

  JOURNAL.Value  := '';
  NATUREPIECE.Value  := '';

  IF S ='Maj' {Outil de Màj} then
  begin
    SetControlVisible('ZZDATEPAQUETMAX', True);
    SetControlVisible('TE_DATEPAQUETMAX', True);
  end;

  DEVISE.Value:=V_PGI.DevisePivot ;

  DATECREATION.Text   := StDate1900 ;
  DATECREATION_.Text  := StDate2099 ;
  DATEECHEANCE.Text   := StDate1900 ;
  DATEECHEANCE_.Text  := StDate2099 ;
  DATEPAQUETMAX.Text  := StDate2099 ;

  {Exigibilité connue dans les paramètres société}
  GENERAL.Plus:='(G_GENERAL LIKE "'+GetParamSocSecur('SO_COLLCLIENC','')
                  +'%" OR G_GENERAL LIKE "'+GetParamSocSecur('SO_COLLFOUENC','')+'%")';

  XX_WHEREINIT.Text:=' (ECRITURE.E_GENERAL LIKE "'+GetParamSocSecur('SO_COLLCLIENC','')
                    +'%" OR ECRITURE.E_GENERAL LIKE "'+GetParamSocSecur('SO_COLLFOUENC','')+'%")';

  If S ='Maj' then
    XX_WHEREINIT.Text:=XX_WHEREINIT.Text
      +' AND J_NATUREJAL="OD"'
      +' AND (ECRITURE.E_MODESAISIE = "" OR ECRITURE.E_MODESAISIE = "-")'
      +' AND (ECRITURE.E_ECHEENC1=0 AND ECRITURE.E_ECHEENC2=0'
      +' AND ECRITURE.E_ECHEENC3=0 AND ECRITURE.E_ECHEENC4=0 AND ECRITURE.E_ECHEDEBIT=0)'
      +' AND (E2.E_LETTRAGE IS NOT NULL AND E2.E_LETTRAGE<>"")'
      +' AND (E2.E_GENERAL LIKE "'+GetParamSocSecur('SO_COLLCLIENC','')+'%"'
      +' OR E2.E_GENERAL LIKE "'+GetParamSocSecur('SO_COLLFOUENC','')+'%")'
      + ' AND ( (ECRITURE.E_ETATLETTRAGE<>"TL" AND ECRITURE.E_ETATLETTRAGE<>"RI")'
      + ' OR (ECRITURE.E_ETATLETTRAGE="TL" AND ECRITURE.E_DATEPAQUETMAX>"'+ USDateTime(StrToDate(DATEPAQUETMAX.Text)) + '"))';

  //XX_WHEREINIT.Text:=XX_WHEREINIT.Text+' AND ECRITURE.E_EDITEETATTVA<>"#" ';

  If VH^.CPExoRef.Code<>'' then
  begin
    EXERCICE.Value      := VH^.CPExoRef.Code ;
    DATECOMPTABLE.Text  := DateToStr(VH^.CPExoRef.Deb) ;
    DATECOMPTABLE_.Text := DateToStr(VH^.CPExoRef.Fin) ;
  end
  else
  begin   // uniquement si exercice ouvert
    If ExisteSQL('SELECT EX_EXERCICE FROM EXERCICE WHERE EX_EXERCICE="' + VH^.Entree.Code
                +'" AND (EX_ETATCPTA="OUV" OR EX_ETATCPTA="CPR")') then
    begin
        EXERCICE.Value      := VH^.Entree.Code ;
        DATECOMPTABLE.Text  := DateToStr(V_PGI.DateEntree) ;
        DATECOMPTABLE_.Text := DateToStr(V_PGI.DateEntree) ;
    end
    else
    begin
        EXERCICE.Value      := VH^.Encours.Code ;
        DATECOMPTABLE.Text  := DateToStr(VH^.Encours.Deb) ;
        DATECOMPTABLE_.Text := DateToStr(VH^.Encours.Deb) ;
    end;

  PositionneEtabUser( ETABLISSEMENT ) ;

  end;


  JOURNAL.Plus:=' AND J_FERME="-"';//vu avec OG le 23/10/2007// AND (J_NATUREJAL="CAI" OR J_NATUREJAL="BQE")';
  //positionnement sur le 1er journal de la liste
  JOURNAL.Value:=JOURNAL.Values[0];

//  XX_WHERE.Text:=StringReplace(XX_WHERE.Text, 'E_', 'ECRITURE.E_'  , [rfReplaceAll]);

//  lStTemp := RecupWhereCritere(TPageControl(GetControl('Pages', True)));


end;

function TOF_CPTVAMODIF.RecupValControl : String;
Var j,i,k : integer ;
    C   : TControl ;
    StWhere : String;
    nom, op, fin, valeur: string;
BEGIN

for j:=0 to Pages.PageCount-1 do
  for i:=0 to Pages.Pages[j].ControlCount-1 do
  BEGIN
  C:=Pages.Pages[j].Controls[i] ;
  if C is TGroupBox then
    BEGIN
      if (C.Visible) and (C.Enabled) then
        for k:=0 to TGroupBox(C).ControlCount-1 do
        begin
         // CP:=TGroupBox(C).Controls[k] ;
        end;
    END
    ELSE
    IF (Copy(C.Name,1,2)='ZZ')
    and (GetControlText(C.Name)<>'')
    and (GetControlText(C.Name)<>StDate1900)
    and (GetControlText(C.Name)<>StDate2099) then
    BEGIN
      nom:=''; op:=''; fin:='"';
      If ((C is THEdit) and (THEdit(C).Operateur=Commence))
      or ((C is THValComboBox) and (THValComboBox(C).Operateur=Commence)) then
      begin
        op:=' LIKE "'; fin:='%"';
      end
      else
        If ((C is THEdit) and (THEdit(C).Operateur=Egal))
        or ((C is THValComboBox) and (THValComboBox(C).Operateur=Egal))      then op:=' ="'
        else
          If ((C is THEdit) and (THEdit(C).Operateur=Superieur))
          or ((C is THValComboBox) and (THValComboBox(C).Operateur=Superieur))  then op:=' >="'
          else
            If ((C is THEdit) and (THEdit(C).Operateur=Inferieur))
            or ((C is THValComboBox) and (THValComboBox(C).Operateur=Inferieur))   then op:=' <="';


       If ((C is THEdit) and (THEdit(C).OpeType=OtDate)) then
          valeur:=USDateTime(StrToDate(GetControlText(C.Name)))
       else
          valeur:=GetControlText(C.Name);

       nom:=StringReplace(C.Name,'_', '',[rfReplaceAll]);          {ZZGENERAL_ ==> ZZGENERAL}
       nom:=StringReplace(nom,'ZZ', 'ECRITURE.E_',[rfReplaceAll]); {ZZGENERAL  ==> ECRITURE.E_GENERAL}

       StWhere:=StWhere+' AND '+nom+op+valeur+fin;

    END ;
  END ;

  Result:=StWhere;

END ;

procedure TOF_CPTVAMODIF.BValideClick(Sender: TObject);
begin
 If (Fliste.NbSelected=0) and (not Fliste.AllSelected) then Exit ;
 OuvrirClick ;
 TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

function TOF_CPTVAMODIF.AddFille (T : TOB ; Q : TQuery) : TOB;
var
      TF : Tob;
      {$IFNDEF EAGLCLIENT}
      i : Integer;
      {$ENDIF}
begin
      TF := Tob.Create('ECRITURE_', T, -1);
      {$IFDEF EAGLCLIENT}
      // Renseigne les valeurs
      TF.Dupliquer(Q.CurrentFille,true,true);
      {$ELSE}
      // Renseigne les valeurs
      for i := 0 to Q.FieldCount - 1 do
        TF.AddChampSupValeur(Q.Fields.Fields[i].FieldName, Q.Fields.Fields[i].AsString, False);
      {$ENDIF}

      Result := TF;
end;

function TOF_CPTVAMODIF.ConstitueLigneOD(TobDetail : Tob) : String;
begin
result:=' WHERE E_JOURNAL="'+TobDetail.GetString('E_JOURNAL')+'"'
       +' AND E_EXERCICE="'+QuelExo(DateToStr(TobDetail.GetDateTime('E_DATECOMPTABLE')))+'"'
       +' AND E_DATECOMPTABLE="'+USDATETIME(TobDetail.GetDateTime('E_DATECOMPTABLE'))+'"'
       +' AND E_NUMLIGNE='+TobDetail.GetString('E_NUMLIGNE')
       +' AND E_NUMECHE='+TobDetail.GetString('E_NUMECHE')
       +' AND E_ECHE="X" AND E_NUMECHE>0 AND E_QUALIFPIECE="N"'
       +' AND E_NUMEROPIECE='+TobDetail.GetString('E_NUMEROPIECE') ;
end;

function TOF_CPTVAMODIF.ConstituePieceOD(TypeEcr : integer ; TobDetail : Tob) : String;
begin
result:=' WHERE E_JOURNAL="'+TobDetail.GetString('E_JOURNAL')+'"'
       +' AND E_EXERCICE="'+QuelExo(DateToStr(TobDetail.GetDateTime('E_DATECOMPTABLE')))+'"'
       +' AND E_DATECOMPTABLE="'+USDATETIME(TobDetail.GetDateTime('E_DATECOMPTABLE'))+'"'
       +' AND E_NUMLIGNE<>'+TobDetail.GetString('E_NUMLIGNE')
       +' AND E_NUMECHE='+TobDetail.GetString('E_NUMECHE')
       +' AND E_ECHE="X" AND E_NUMECHE>0 AND E_QUALIFPIECE="N"'
       +' AND E_NUMEROPIECE='+TobDetail.GetString('E_NUMEROPIECE')
       +' AND E_LETTRAGE<>"" AND (E_ETATLETTRAGE="TL" OR E_ETATLETTRAGE="PL")' ;

if typeEcr=2 then
  result:=result+' AND E_AUXILIAIRE="'+TobDetail.GetString('E_AUXILIAIRE')+'"' ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 29/10/2007
Modifié le ... :   /  /
Description .. : Recherche du sens de la recopie des infos : en sens
Suite ........ : inverse ou pas
Mots clefs ... :
*****************************************************************}
function TOF_CPTVAMODIF.SensInverse(TobDet : Tob) : Boolean;
var
  QLig, Qsens : TQuery;
  SQL, StLig: String;
begin
  Result:=false;

  {Recherche d'une ligne dans la pièce avec un collectif différent de celui de la ligne sélectionnée}

  StLig:=ConstitueLigneOD(TobDet);
  SQL:='SELECT E_JOURNAL, E_DATECOMPTABLE, E_EXERCICE, E_NUMLIGNE, E_NUMECHE, E_NUMEROPIECE, E_GENERAL FROM ECRITURE '+StLig ;
  QLig:=OpenSQL(SQL,True) ;

  If not QLig.EOF then
  begin

      SQL:='SELECT E_GENERAL FROM ECRITURE '
           +' WHERE E_JOURNAL="'+Qlig.FindField('E_JOURNAL').AsString+'"'
           +' AND E_EXERCICE="'+QuelExo(DateToStr(Qlig.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
           +' AND E_DATECOMPTABLE="'+USDATETIME(Qlig.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
           +' AND E_NUMLIGNE<>'+Qlig.FindField('E_NUMLIGNE').AsString
           +' AND E_NUMECHE='+Qlig.FindField('E_NUMECHE').AsString
           +' AND E_ECHE="X" AND E_QUALIFPIECE="N"'
           +' AND E_NUMEROPIECE='+Qlig.FindField('E_NUMEROPIECE').AsString
           +' AND ((E_GENERAL LIKE "'+GetParamSocSecur('SO_COLLCLIENC','')
           +'%" AND "'+Qlig.FindField('E_GENERAL').AsString+'" LIKE "'+GetParamSocSecur('SO_COLLFOUENC','')+'%")'
           +' OR (E_GENERAL LIKE "'+GetParamSocSecur('SO_COLLFOUENC','')
           +'%" AND "'+Qlig.FindField('E_GENERAL').AsString+'" LIKE "'+GetParamSocSecur('SO_COLLCLIENC','')+'%"))';

      Qsens:=OpenSQL(SQL,True) ;
      if Not Qsens.EOF then Result:=true;

  end;

  Ferme(Qsens);
  Ferme(Qlig);

   {Passage en revue des lignes de la pièce à la recherche de la nature des comptes
   for i := 0 to T.Detail.Count -1 do
   begin
       If T.Detail[i].GetString
   end;   }
end;


function TOF_CPTVAMODIF.ConstituePieceOrig(TypeEcr : integer ; Q : TQuery) : String;
begin
result:=' WHERE E_ETATLETTRAGE="'+Q.FindField('E_ETATLETTRAGE').AsString+'"'
       +' AND E_LETTRAGE="'+Q.FindField('E_LETTRAGE').AsString+'"'
       +' AND E_GENERAL="'+Q.FindField('E_GENERAL').AsString+'"'
       +' AND E_AUXILIAIRE="'+Q.FindField('E_AUXILIAIRE').AsString+'"'
       +' AND E_NUMEROPIECE<>'+Q.FindField('E_NUMEROPIECE').AsString;

if typeEcr=1 then
  result:=result+' AND E_TIERSPAYEUR<>"" AND E_PIECETP<>""';
end;

procedure TOF_CPTVAMODIF.OuvrirClick;
var
    lMulQ, Q1 : TQuery ;
    i : integer ;
    TobEcr : TOB ;
    OM : TMod;
    Exo, StW, SQL, Regime, Nature, CodeTva, Exige:string;
    Client : boolean;
    Montant, TauxTva, Valeur : double;
    RTVA   : Enr_Base ;
//    TInfo : TOB;
    CodeEcran, NumTaux : String;
    Coef : ShortInt;
begin

If Maj.Text='X' then
  OutilMaj
else
BEGIN

{$IFDEF EAGLCLIENT}
  lMulQ := FEcran.Q.TQ;
  LMulQ.Seek( FListe.Row - 1 );
  LMulQ.PutGridDetail(FListe,true,true,''); // Affichage des en-têtes
{$ELSE}
  lMulQ := FEcran.Q;
{$ENDIF}

  TobEcr := Tob.Create('ECRITURE_', nil, -1) ;

  If TFmul(Ecran).FListe.AllSelected then
  begin
      LMulQ.First;

      While not LMulQ.EOF do
      begin
         AddFille(TobEcr, LMulQ);
         LMulQ.Next;
      end;
  end
  else
  if FListe.NbSelected = 0 then
  begin
    MoveCur(FALSE);

    {$IFDEF EAGLCLIENT}
    LMulQ.Seek(FListe.Row - 1);
    {$ENDIF}

    AddFille(TobEcr, LMulQ);
  end
  else
  for i := 0 to FListe.NbSelected - 1 do
  begin
    MoveCur(FALSE);
    FListe.GotoLeBookMark(i);
    {$IFDEF EAGLCLIENT}
    LMulQ.Seek(FListe.Row - 1);
    {$ENDIF}

    AddFille(TobEcr, LMulQ);
  end;


IF TobEcr.Detail.Count = 1 THEN {Traitement manuel : au plus un enreg sélectionné}
BEGIN

  if ((LMulQ.EOF) And (LMulQ.Bof)) then Exit ;
  OM:=Nil ; UpSQL:='' ;
  StW:=ConstitueLigneOD(TobEcr.Detail[0]);

  SQL:='SELECT E_DEBIT,E_CREDIT,E_REGIMETVA,E_NATUREPIECE,E_TVA,E_DATEMODIF,E_GENERAL,E_AUXILIAIRE,'
      +'E_DEBITDEV,E_CREDITDEV,E_DATECOMPTABLE,E_DATEREFEXTERNE,E_MODEPAIE,E_DATEECHEANCE,G_SENS,'
      +'E_ECHEENC1,E_ECHEENC2,E_ECHEENC3,E_ECHEENC4,E_ECHEDEBIT '
      +'FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL'
      +StW;

  Q1:=OpenSQL(SQL,True) ; Client:=False ;
  if Not Q1.EOF then
  BEGIN
       Regime:=Q1.FindField('E_REGIMETVA').AsString ;
       Nature:=Q1.FindField('E_NATUREPIECE').AsString ;
       CodeTva:=Q1.FindField('E_TVA').AsString ;
       Client:=EstCeClient(Q1.FindField('E_GENERAL').AsString,Q1.FindField('E_AUXILIAIRE').AsString) ;
       OM:=TMOD.Create ;
       RemplirOM(OM,Q1) ;
  END ;

  Ferme(Q1) ;
  if OM<>Nil then
  BEGIN
     RTVA.Regime:=Regime ; RTVA.Client:=Client ; RTVA.Action:=taModif ;
     RTVA.Nature:=Nature ; RTVA.CodeTva:=CodeTva ;
     if SaisieBasesHT(OM,RTVA) then
     BEGIN
        if OM.MODR.ModifTva then
        BEGIN
           UpSQL:='UPDATE ECRITURE SET E_ECHEENC1='+StrfPoint(OM.MODR.TabEche[1].TAV[1])
                 +', E_ECHEENC2='+StrfPoint(OM.MODR.TabEche[1].TAV[2])
                 +', E_ECHEENC3='+StrfPoint(OM.MODR.TabEche[1].TAV[3])
                 +', E_ECHEENC4='+StrfPoint(OM.MODR.TabEche[1].TAV[4])
                 +', E_ECHEDEBIT='+StrfPoint(OM.MODR.TabEche[1].TAV[5])
                 +', E_TVA="'+RTVA.CodeTva+'"'
                 +', E_DATEMODIF="'+USTime(NowFutur)+'"'
                 +StW ;
           if Transactions(MajModTvaEnc,3)<>oeOk then MessageAlerte('ATTENTION. Modification non enregistrée !') ;
        END ;
     END ;
  END ;

  OM.Free ;

END
ELSE
BEGIN {Traitement automatique si plus d'un enreg sélectionné}

    {Recherche infos communes}
    for i := 0 to TobEcr.Detail.count-1 do
    begin
      MoveCur(FALSE);

      UpSQL:='SELECT T_REGIMETVA, T_TVAENCAISSEMENT FROM TIERS '
            +'WHERE T_AUXILIAIRE="'+TobEcr.Detail[i].GetString('E_AUXILIAIRE')+'"';
      Q1:=OpenSQL(UpSQL,True) ;

      if Not Q1.EOF then
      begin
        If Regime='' then Regime:=Q1.FindField('T_REGIMETVA').AsString ;

        If Exige ='' then Exige:=Q1.FindField('T_TVAENCAISSEMENT').AsString;

        If Q1.FindField('T_REGIMETVA').AsString <> Regime then
          Regime:='X'; {Infos identiques affichées}

        If Q1.FindField('T_TVAENCAISSEMENT').AsString <> Exige then
          Exige:='X';

      end;

      Ferme(Q1) ;
    end;

    {Affichage infos communes + choix du taux à appliquer automatiquement}
    If CPLanceFiche_ParamTVA(Regime+';'+Exige) <> 'S' then
    BEGIN

//        TInfo := TOB.Create('', nil, -1);
//        TInfo := TheTob;

        If TheTob<>nil then
        begin
          NumTaux:= TheTob.GetString('NUMTAUX');
          CodeEcran:= TheTob.GetString('CODETVA');
        end;

        for i := 0 to TobEcr.Detail.count - 1 do
        begin
          exo:=QuelExo(DateToStr(TobEcr.Detail[i].GetDateTime('E_DATECOMPTABLE')));
          UpSQL:='' ;
          StW:=ConstitueLigneOD(TobEcr.Detail[i]);

          SQL:='SELECT E_DEBIT,E_CREDIT,E_REGIMETVA,E_NATUREPIECE,E_TVA,E_DATEMODIF,E_GENERAL,E_AUXILIAIRE,'
          +'E_DEBITDEV,E_CREDITDEV,E_DATECOMPTABLE,E_DATEREFEXTERNE,E_MODEPAIE,E_DATEECHEANCE,G_SENS,'
          +'E_ECHEENC1,E_ECHEENC2,E_ECHEENC3,E_ECHEENC4,E_ECHEDEBIT,E_TVA '
          +'FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL'+StW ;

          Q1:=OpenSQL(SQL,True) ;
          Client:=False ;
          Montant:=0;
          Coef:=1;
          if Not Q1.EOF then
          BEGIN
               Montant:=Q1.FindField('E_DEBIT').AsFloat+Q1.FindField('E_CREDIT').AsFloat ;
               Regime:=Q1.FindField('E_REGIMETVA').AsString ;
               Nature:=Q1.FindField('E_NATUREPIECE').AsString ;
               CodeTva:=Q1.FindField('E_TVA').AsString ;
               Client:=EstCeClient(Q1.FindField('E_GENERAL').AsString,Q1.FindField('E_AUXILIAIRE').AsString) ;

               If ((Q1.FindField('G_SENS').AsString='D') And (Q1.FindField('E_CREDIT').AsFloat>0))
               Or ((Q1.FindField('G_SENS').AsString='C') And (Q1.FindField('E_DEBIT').AsFloat >0)) then
                  Coef:=-1; {Ecriture en sens inverse du sens du compte général}

          END ;
          Ferme(Q1) ;

          UpSQL:='SELECT T_REGIMETVA, T_TVAENCAISSEMENT FROM TIERS '
                +'WHERE T_AUXILIAIRE="'+TobEcr.Detail[i].GetString('E_AUXILIAIRE')+'"';
          Q1:=OpenSQL(UpSQL,True) ;

          if Not Q1.EOF then
          begin
            Regime:=Q1.FindField('T_REGIMETVA').AsString ;
            If Client then
              Exige :=GetParamSocSecur('SO_TVAENCAISSEMENT','')
            else
              Exige :=Q1.FindField('T_TVAENCAISSEMENT').AsString;
          end;

          Ferme(Q1) ;

          {RAZ avant Maj}
          UpSql:='UPDATE ECRITURE SET E_ECHEENC1=0, E_ECHEENC2=0, E_ECHEENC3=0,'
                +'E_ECHEENC4=0, E_ECHEDEBIT=0, E_DATEMODIF="'+USTime(NowFutur)+'"'
                +StW;

          if Transactions(MajModTvaEnc,3)<>oeOk then MessageAlerte('Attention : modification non enregistrée !') ;


          UpSql:=' SELECT TV_TAUXACH, TV_TAUXVTE FROM TXCPTTVA'
                  +' WHERE TV_TVAOUTPF = "TX1"'
                  +' AND TV_REGIME = "'+Regime+'"'
                  +' AND TV_CODETAUX = "'+CodeEcran+'"';

          Q1:=OpenSQL(UpSQL,True) ;

          TauxTva:=0;
          if Not Q1.EOF then
          begin
              If Client then
                TauxTva:=Q1.FindField('TV_TAUXVTE').AsFloat
              else
                TauxTva:=Q1.FindField('TV_TAUXACH').AsFloat;
          end;

          Ferme(Q1) ;

          Valeur:=Coef*Arrondi(Montant /(1.0+(TauxTva/100.0)), V_PGI.OkDecV) ;
          {FQ22230  23.01.08  YMO On prend le taux 1 pour ECHEDEBIT}
          If NumTaux<>'0' then
            UpSQL:='UPDATE ECRITURE SET E_ECHEENC'+NumTaux+'='+StrFPoint(Valeur)
                  +', E_TVA="'+CodeEcran+'"'
                  +', E_DATEMODIF="'+USTime(NowFutur)+'"'
                  +StW
          else
             UpSQL:='UPDATE ECRITURE SET E_ECHEDEBIT='+StrFPoint(Valeur)
                   +', E_DATEMODIF="'+USTime(NowFutur)+'"'
                   +StW;

          if Transactions(MajModTvaEnc,3)<>oeOk then MessageAlerte('Attention : Modification non enregistrée !') ;

        end;

    END;

END ;

if Not FListe.AllSelected then Fliste.ClearSelected else FListe.AllSelected:=False ;

TobEcr.free;

END;

end;

procedure TOF_CPTVAMODIF.CopieEche(TypeEcr : integer ; TobDet : Tob);
var
    QPie, QOri : TQuery ;
    SQL, CodTva, RegTva : String ;
    Stlig, StPie, StOri : String ;
    k : integer ;
    Eche : Array[1..5] of Double;
    Coef : ShortInt;
    OkModif : boolean ;
begin

  StLig:=ConstitueLigneOD(TobDet);
  StPie:=ConstituePieceOD(TypeEcr, TobDet);

  SQL:='SELECT E_NUMEROPIECE, E_GENERAL, E_AUXILIAIRE, E_LETTRAGE, E_ETATLETTRAGE, E_TIERSPAYEUR FROM ECRITURE '+StPie ;

  QPie:=OpenSQL(SQL,True) ;

  While not QPie.EOF do
  begin
      StOri:=ConstituePieceOrig(TypeEcr, QPie);

      SQL:='SELECT E_ECHEENC1,E_ECHEENC2,E_ECHEENC3,E_ECHEENC4,E_ECHEDEBIT, E_REGIMETVA, E_TVA FROM ECRITURE '+StOri ;

      QOri:=OpenSQL(SQL,True) ;

      for k := 1 to 5 do Eche[k]:=0;
      RegTva:='';
      CodTva:='';

      While not QOri.EOF do
      begin
        for k := 1 to 4 do Eche[k]:=Eche[k]+QOri.FindField('E_ECHEENC'+inttostr(k)).AsFloat;
        Eche[5]:=Eche[5]+QOri.FindField('E_ECHEDEBIT').AsFloat;
        RegTva:=QOri.FindField('E_REGIMETVA').AsString; {FQ22036  13.12.2007  YMO}
        CodTva:=QOri.FindField('E_TVA').AsString;
        QOri.Next;
      end;

      Ferme(QOri) ;

      QPie.Next;
  end;

  Ferme(QPie) ;

  If (TypeEcr=3) and (SensInverse(TobDet)) then Coef:=-1 else Coef:=1;

  OkModif:=False;
  for k := 1 to 5 do
    if Eche[k]<>0 then OkModif:=true;

  If OkModif then
    UpSQL:='UPDATE ECRITURE SET E_ECHEENC1='+StrFPoint(Arrondi(Coef*Eche[1],V_PGI.OkDecV))
     +', E_ECHEENC2='+StrFPoint(Arrondi(Coef*Eche[2],V_PGI.OkDecV))
     +', E_ECHEENC3='+StrFPoint(Arrondi(Coef*Eche[3],V_PGI.OkDecV))
     +', E_ECHEENC4='+StrFPoint(Arrondi(Coef*Eche[4],V_PGI.OkDecV))
     +', E_ECHEDEBIT='+StrFPoint(Arrondi(Coef*Eche[5],V_PGI.OkDecV))
     +', E_TVA="'+CodTva+'"'
     +', E_REGIMETVA="'+RegTva+'"'
     +', E_DATEMODIF="'+USTime(NowFutur)+'"'
     +StLig;

  ExecuteSql(UpSQL);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 29/10/2007
Modifié le ... :   /  /
Description .. : Outil de mise à jour des bases de TVA
Mots clefs ... : OUTIL MAJ
*****************************************************************}
procedure TOF_CPTVAMODIF.OutilMaj;
var
    lMulQ : TQuery ;
    i : integer ;
    TobEcr : TOB ;
    NbD : integer;
begin

{$IFDEF EAGLCLIENT}
  lMulQ := FEcran.Q.TQ;
  LMulQ.Seek( FListe.Row - 1 );
  LMulQ.PutGridDetail(FListe,true,true,''); // Affichage des en-têtes
{$ELSE}
  lMulQ := FEcran.Q;
{$ENDIF}

  TobEcr := Tob.Create('ECRITURE_', nil, -1) ;

  NbD     := FListe.NbSelected ;
  InitMove(NbD,'') ;
  If TFmul(Ecran).FListe.AllSelected then
  begin
      LMulQ.First;

      While not LMulQ.EOF do
      begin
         AddFille(TobEcr, LMulQ);
         LMulQ.Next;
      end;
  end
  else
  if FListe.NbSelected = 0 then
  begin
    MoveCur(FALSE);

    {$IFDEF EAGLCLIENT}
    LMulQ.Seek(FListe.Row - 1);
    {$ENDIF}

    AddFille(TobEcr, LMulQ);
  end
  else
  for i := 0 to FListe.NbSelected - 1 do
  begin
    MoveCur(FALSE);
    FListe.GotoLeBookMark(i);
    {$IFDEF EAGLCLIENT}
    LMulQ.Seek(FListe.Row - 1);
    {$ENDIF}

    AddFille(TobEcr, LMulQ);
  end;
  FiniMove ;

    {StLig : reprise de l'écriture su Mul sue laquelle on va recopier les infos ECHEENC
     StPie : ligne lettrée correspondante dans la même pièce
     StOri : ligne d'origine sur laquelle sonr les infos ECHEENC}

    for i := 0 to TobEcr.Detail.Count - 1 do
    begin
      {OD de Tiers payeur : 1er cas }
      If (TobEcr.Detail[i].GetString('E_QUALIFORIGINE')='TP') then
        CopieEche(1, TobEcr.Detail[i])
      else
      {OD de suivi d'effets}
      If ExisteSQL('SELECT 1 FROM JOURNAL WHERE J_JOURNAL="'+TobEcr.Detail[i].GetString('E_JOURNAL')+'" AND J_EFFET="X"') then
        CopieEche(2, TobEcr.Detail[i])
      else
      {autres types d'OD}
        CopieEche(3, TobEcr.Detail[i]);

    end;

    TobEcr.free;

    if Not FListe.AllSelected then Fliste.ClearSelected else FListe.AllSelected:=False ;

END ;

procedure TOF_CPTVAMODIF.FListeDblClick(Sender: TObject);
begin
  If Maj.Text='X' then
    BZoomPieceClick(Sender)
  else
  begin
    OuvrirClick ;
    TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
  end;

end;

procedure TOF_CPTVAMODIF.BZoomPieceClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  TrouveEtLanceSaisie(TFMul(Ecran).Q.TQ,taConsult,GetControlText('ZZQUALIFPIECE'));
{$ELSE}
  TrouveEtLanceSaisie(TFMul(Ecran).Q,taConsult,GetControlText('ZZQUALIFPIECE'));
{$ENDIF}
end;

procedure TOF_CPTVAMODIF.EXERCICEChange(Sender: TObject);
begin
 	ExoToDates( GetControlText('ZZEXERCICE'), TEdit(GetControl('ZZDATECOMPTABLE', True)), TEdit(GetControl('ZZDATECOMPTABLE_', True)) ) ;
	if ((GetControlText('ZZEXERCICE')='')) then
  BEGIN
    SetControlText('ZZDATECOMPTABLE',		stDate1900) ;
    SetControlText('ZZDATECOMPTABLE_',	stDate2099) ;
  END ;
end;

procedure TOF_CPTVAMODIF.InitMsgBox;
begin
  MsgBox := THMsgBox.create(FMenuG);
  MsgBox.Mess.Add('0;Validation;Ce traitement va valider toutes les pièces sélectionnées. Désirez-vous continuer?;Q;YN;Y;;');
  MsgBox.Mess.Add('1;Dévalidation;Ce traitement va dévalider toutes les pièces sélectionnées. Désirez-vous continuer?;Q;YN;Y;;');
end;

Function TOF_CPTVAMODIF.EstCeClient ( Gene,Auxi : String ) : boolean ;
Var Q2 : TQuery ;
BEGIN
Result:=False ;
if Auxi<>'' then
BEGIN
   Q2:=OpenSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="'+Auxi+'" AND (T_NATUREAUXI="CLI" OR T_NATUREAUXI="AUD")',True) ;
   if Not Q2.EOF then Result:=True ;
   Ferme(Q2) ;
END else
BEGIN
   Q2:=OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+Gene+'" AND G_NATUREGENE="TID"',True) ;
   if Not Q2.EOF then Result:=True ;
   Ferme(Q2) ;
END ;

END ;

procedure TOF_CPTVAMODIF.RemplirOM ( OM : TMOD ; Q1 : TQuery ) ;
Var
  k : integer ;
BEGIN
FillChar(OM.MODR,Sizeof(OM.MODR),#0) ;
OM.ModR.Action:=taModif ;
OM.ModR.NbEche:=1 ;
OM.ModR.TotalAPayerP:=Q1.FindField('E_DEBIT').AsFloat+Q1.FindField('E_CREDIT').AsFloat ;
If (Q1.FindField('G_SENS').AsString='D') And (Q1.FindField('E_CREDIT').AsFloat<>0) then
//    OM.ModR.TabEche[1].Sens:='-'; {sens opposé}
  OM.ModR.TotalAPayerP:=-OM.ModR.TotalAPayerP;

OM.ModR.TotalAPayerD:=Q1.FindField('E_DEBITDEV').AsFloat+Q1.FindField('E_CREDITDEV').AsFloat ;
OM.ModR.CodeDevise:=V_PGI.DevisePivot ;
OM.ModR.TauxDevise:=1.0 ;
OM.ModR.Quotite:=1 ;
OM.ModR.Decimale:=V_PGI.OkDecV ;
OM.ModR.DateFact:=Q1.FindField('E_DATECOMPTABLE').AsDateTime ;
OM.ModR.DateBL:=OM.MODR.DateFact ;
// GP REGL
OM.ModR.DateFactExt:=Q1.FindField('E_DATEREFEXTERNE').AsDateTime ;
If OM.ModR.DateFactExt=IDate1900 Then OM.ModR.DateFactExt:=OM.ModR.DateFact ;
OM.ModR.Aux:=Q1.FindField('E_AUXILIAIRE').AsString ;
OM.ModR.ModifTva:=False ;
OM.ModR.TabEche[1].ModePaie:=Q1.FindField('E_MODEPAIE').AsString ;
OM.ModR.TabEche[1].DateEche:=Q1.FindField('E_DATEECHEANCE').AsDateTime ;
OM.ModR.TabEche[1].MontantP:=OM.MODR.TotalAPayerP ;
OM.ModR.TabEche[1].MontantD:=OM.MODR.TotalAPayerD ;
OM.ModR.TabEche[1].Pourc:=100.0 ;
{#TAVENC}
for k:=1 to 4 do OM.MODR.TabEche[1].TAV[k]:=Q1.FindField('E_ECHEENC'+IntToStr(k)).AsFloat ;
OM.ModR.TabEche[1].TAV[5]:=Q1.FindField('E_ECHEDEBIT').AsFloat ;

END ;

procedure TOF_CPTVAMODIF.MajModTvaEnc ;
BEGIN
if ExecuteSQL(UpSQL)<>1 then V_PGI.IoError:=oeUnknown ;
END ;

procedure TOF_CPTVAMODIF.ValideClick;
begin
end;


procedure TOF_CPTVAMODIF.OnClose ;
begin
  Inherited ;
end ;


Initialization
  registerclasses ( [ TOF_CPTVAMODIF ] ) ;
end.
