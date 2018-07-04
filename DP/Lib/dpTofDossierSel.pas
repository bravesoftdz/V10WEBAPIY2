{***********UNITE*************************************************
Auteur  ...... : MD
Créé le ...... : 30/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YYDOSSIER_SEL ()
Mots clefs ... : TOF;DOSSIERSEL
*****************************************************************}
Unit dpTofDossierSel ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     eMul, MaineAGL,
{$ELSE}
     Mul, Fe_Main,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HDB,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HTB97;

Type
  TOF_DOSSIERSEL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure BCherche_OnClick    (Sender: TObject);
    procedure FListe_OnDblClick   (Sender: TObject);
    procedure BAnnuaire_OnClick   (Sender: TObject);
    procedure SANSGRPCONF_OnClick (Sender: TObject);
  private
    NoDossier    : String;
    bOnlyCabinet : Boolean;
    procValideDecla : TNotifyEvent ;//LMO
    FiltreDonnees : string ;//LM20071008
    procedure bValider_OnClick(sender:TObject) ;//LMO
    function  getInfoRetour : string ;//LMO
  end ;


// $$$ JP 06/04/2004: function de lancement fiche de sélection d'un dossier
function DP_SelectUnDossier (strDossier:string; bDossierCabinet:boolean):string;


//////////////// IMPLEMENTATION //////////////////
Implementation

uses
    galOutil,

    {$IFDEF VER150}
    Variants,
    {$ENDIF}

    UtilDossier; // $$$ JP 26/04/06

// $$$ JP 06/04/2004: function de lancement fiche de sélection d'un dossier
function DP_SelectUnDossier (strDossier:string; bDossierCabinet:boolean):string;
begin
     if bDossierCabinet = TRUE then
         Result := AGLLanceFiche ('YY', 'YYDOSSIER_SEL', '', '', strDossier + ';' + 'ONLYCAB')
     else
         Result := AGLLanceFiche ('YY', 'YYDOSSIER_SEL', '', '', strDossier);
end;


procedure TOF_DOSSIERSEL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_DOSSIERSEL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_DOSSIERSEL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_DOSSIERSEL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_DOSSIERSEL.OnArgument (S : String ) ;
var ChXXWhere    : String;
    ctrl : TControl ;
begin
  Inherited ;

  //+LMO
  if (Pos ('MULTISELECT', s) > 0) then
  begin
  {$IFNDEF EAGLCLIENT} //+LM20060901
    ThDbGrid(GetControl('FListe')).MultiSelection := true ;
  {$ELSE}
    ThGrid(GetControl('FListe')).MultiSelect := true ;
  {$ENDIF}             //-LM20060901

    SetControlVisible('BSELECTALL', true);
  end ;


  ctrl := GetControl('BOUVRIR');
  if (ctrl<>nil) and assigned(TToolBarButton97(Ctrl).OnClick) then
  begin
    procValideDecla := TToolBarButton97(ctrl).OnClick ;
    TToolBarButton97(ctrl).OnClick := bValider_OnClick;
  end ;
  //-LMO

  // Arguments
  if pos ('FILTREDONNEES=', s) >0 then //LM20071008
  begin
    FiltreDonnees:=RecupEtSupprimeArgument('FILTREDONNEES', s) ;
    setcontrolProperty('GROUPECONF', 'DataType', 'YYFILTREDONNEES') ;
  end ;
  NoDossier := ReadTokenSt(S);

  // $$$ JP 06/04/2004 - si sélection dossier du cabinet seulement
  bOnlyCabinet := Pos ('ONLYCAB', S) > 0;

  InitialiserComboGroupeConf (THMultiValCombobox(GetControl('GROUPECONF')), FiltreDonnees); //LM20071008
  ChXXWhere:=GererCritereGroupeConf (THMultiValCombobox(GetControl('GROUPECONF')),
                                     TcheckBox (GetControl ('SANSGRPCONF')).Checked,
                                     false,
                                     FiltreDonnees); //LM20071008
  ChXXWhere:=ChXXWhere+GererCritereDivers (False, False, False, BOnlyCabinet);
  // $$$ JP 04/09/06 - plus autorisé en D7/Unicode ... TEdit(GetControl('XX_WHERE')).Text:=ChXXWhere;
  SetControlText ('XX_WHERE', chXXWhere);

  TToolBarButton97(GetControl('BCherche')).OnClick := BCherche_OnClick;
  THDBGrid(GetControl('FLISTE')).OnDblClick        := FListe_OnDblClick;
  TCheckbox(GetControl('SANSGRPCONF')).OnClick  := SANSGRPCONF_OnClick;

  if GetControl('BAnnuaire')<>Nil then
    TToolBarButton97(GetControl('BAnnuaire')).OnClick := BAnnuaire_OnClick;
end ;

procedure TOF_DOSSIERSEL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_DOSSIERSEL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_DOSSIERSEL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_DOSSIERSEL.BCherche_OnClick(Sender: TObject);
var ChXXWhere  :string;
begin
 NextPrevControl(Ecran, True, True); // FQ 11630

 ChXXWhere:=GererCritereGroupeConf (THMultiValCombobox(GetControl('GROUPECONF')),
                                    TcheckBox (GetControl ('SANSGRPCONF')).Checked,
                                    False,
                                    FiltreDonnees);//LM20071008
 ChXXWhere:=ChXXWhere+GererCritereDivers (False, False, False, BOnlyCabinet);
 // $$$ JP 04/09/06 - plus autorisé en D7/Unicode ... TEdit(GetControl('XX_WHERE')).Text:=ChXXWhere;
 SetControlText ('XX_WHERE', chXXWhere);

 TFMul(Ecran).BChercheClick(Sender);
{$IFDEF EAGLCLIENT}
    // #### A REFAIRE eAGL ####
{$ELSE}
  if NoDossier<>'' then
    TFMul(Ecran).Q.Locate('DOS_NODOSSIER', NoDossier, []);
{$ENDIF}
end;

procedure TOF_DOSSIERSEL.FListe_OnDblClick(Sender: TObject);
var
   strRetour   :string;
begin
     // $$$ JP 26/04/06 - ANN_NOM1 par forcément présent dans le paramétrage de la liste: on le lit si pas présent
     if Not VarIsNull(GetField ('DOS_GUIDPER')) then
     begin
      {$IFNDEF EAGLCLIENT} //+LM20060831
       if ThDbGrid(GetControl('FListe')).MultiSelection then   //+LMO
         ThDbGrid(GetControl('FListe')).FlipSelection
       else
       {$ELSE} //+LM20060901
       if ThGrid(GetControl('FListe')).MultiSelect then
         ThGrid(GetControl('FListe')).FlipSelection(ThGrid(GetControl('FListe')).Row)
       else    //-LM20060901
       {$ENDIF} //+LM20060831
       begin //-LMO
          strRetour := GetField('DOS_NODOSSIER')+';'+GetField('DOS_GUIDPER');
          if VarIsNull (GetField ('ANN_NOM1')) = FALSE then
              strRetour := strRetour + ';' + GetField ('ANN_NOM1')
          else
              strRetour := strRetour + ';' + GetNomDossier (GetField ('DOS_NODOSSIER'));
          TFMul (Ecran).Retour := strRetour;
       end ;
     end;

     Ecran.Close;
end;

procedure TOF_DOSSIERSEL.BAnnuaire_OnClick(Sender: TObject);
begin
  if Not VarIsNull(GetField('DOS_GUIDPER')) and (GetField('DOS_GUIDPER')>'') then
    AGLLanceFiche('YY','ANNUAIRE','', GetField('DOS_GUIDPER'),
     'ACTION=MODIFICATION;;;'+'DOS')  ;
end;

procedure TOF_DOSSIERSEL.SANSGRPCONF_OnClick(Sender: TObject);
begin
  GereCheckboxSansGrpConf(TCheckbox(GetControl('SANSGRPCONF')), THMultiValCombobox(GetControl('GROUPECONF')) );
end;

function TOF_DOSSIERSEL.getInfoRetour : string ;//LMO
var st : string ;
begin
  st := GetField('DOS_NODOSSIER')+';'+
        GetField('DOS_GUIDPER') + ';';
  if not VarIsNull (GetField ('ANN_NOM1')) then st := st + GetField ('ANN_NOM1')
                                           else st := st + GetNomDossier (GetField ('DOS_NODOSSIER'));
  result := st ;
end ;

procedure TOF_DOSSIERSEL.bValider_OnClick(sender:TObject) ;//LMO
var
  {$IFDEF EAGLCLIENT}
    DbG : ThGrid ;
  {$ELSE}
    DbG : ThDbGrid ;
  {$ENDIF}
    i : integer ;
    st : string ;

      procedure _bValider_OnClick ;
      begin
        if st<>'' then st:=st + '|' ;
        st:=st + getInfoRetour ;
      end ;

begin
{$IFDEF EAGLCLIENT} //+LM20060901 remplace les anciennes directives
  DbG := ThGrid(GetControl('FListe')) ;
  if DbG.MultiSelect then
{$ELSE}
  DbG := ThDbGrid(GetControl('FListe')) ;
  if DbG.MultiSelection then
{$ENDIF}
  //if DbG.MultiSelection then
  //-LM20060901
  begin
    st := '' ;
    if (DbG.AllSelected) then
    begin
      with TFMul(Ecran).Q do
      begin
        First ;
        while not EOF do
        begin
          _bValider_OnClick ;
          next ;
        end ;
      end ;
    end
    else if (DbG.nbSelected>0) then
    begin
      for i:= 0 to DbG.nbSelected-1 do
      begin
        DbG.GotoLeBookmark(i) ;
        _bValider_OnClick ;
      end ;
    end ;
    TFMul (Ecran).Retour := st;
  end ;

  procValideDecla(sender) ;
end ;

Initialization
  registerclasses ( [ TOF_DOSSIERSEL ] ) ;
end.

