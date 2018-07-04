unit UtofPtfPiece;

interface
uses  M3FP, StdCtrls,Controls,Classes, Graphics, forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOB,uTofAfBaseCodeAffaire,EntGC,

{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      Fe_Main,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}MajTable,
{$ENDIF}
      UTOF, AglInit, Agent, Ent1, UtilGc,HTB97, QRS1;

Type
     TOF_PtfPiece = Class (TOF_AFBASECODEAFFAIRE)
       private

       //FV1 : 14/03/2017 - FS#2375 - LEHODEY TP - Problème édition portefeuille avec article contenant le symbole _
       CodeArticle  : THEdit;
       CodeArticle1 : THEdit;

       procedure ControleChamp(Champ, Valeur: String);
       procedure GestionFS1635(Lib: String; Affiche: Boolean);

       //FV1 : 14/03/2017 - FS#2375 - LEHODEY TP - Problème édition portefeuille avec article contenant le symbole _
       procedure OnExitArticle(Sender: Tobject);
       
       //uniquement en line
//     procedure DefiniRuptureS1;

       public
       procedure OnArgument (stArgument : String )  ; override ;
       procedure OnUpdate ; override ;
       procedure OnLoad ; override ;
       procedure NomsChampsAffaire( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ; override ;
     END ;

Const
	// libellés des messages
	TexteMessage: array[1..1] of string 	= (
          {1}         'Vous devez sélectionner une nature de pièce.'
                 );
implementation

/////////////////////////////////////////////////////////////////////////////
procedure TOF_PtfPiece.OnArgument (stArgument : String ) ;
Var CC : THValComboBox ;
    x				: integer;
    Critere : String;
    ValMul	: String;
    ChampMul: String;
begin
    Inherited;

    Setcontroltext('XX_WHERE', '');

{$IFDEF BTP}
  Repeat

  Critere:=uppercase(ReadTokenSt(stArgument)) ;
  if Critere<>'' then
      begin
      x:=pos('=',Critere);
      if x<>0 then
         begin
         ChampMul:=copy(Critere,1,x-1);
         ValMul:=copy(Critere,x+1,length(Critere));
         end
      else
         ChampMul := Critere;
      ControleChamp(ChampMul, ValMul);
      end;
  until Critere='';

  //FV1 : 14/03/2017 - FS#2375 - LEHODEY TP - Problème édition portefeuille avec article contenant le symbole _
  CodeArticle   := THEdit(GetControl('CODEARTICLE'));
  if CodeArticle <> nil then CodeArticle.OnExit := OnExitArticle;
  //
  CodeArticle1  := THEdit(GetControl('CODEARTICLE1'));

{$ELSE}
  if stArgument='VENTE' then St_Plus:='AND GPP_VENTEACHAT="VEN"'
  else if stArgument='ACHAT' then St_Plus:='AND GPP_VENTEACHAT="ACH"'
  else St_Plus:='AND GPP_VENTEACHAT="AUT"' ;
  {$IFDEF NOMADE}
  if stArgument = 'ACHAT'then
    St_Plus := St_Plus + ' AND (GPP_NATUREPIECEG="CF" OR GPP_TRSFACHAT="X")'
  else
    St_Plus := St_Plus + ' AND (GPP_NATUREPIECEG IN ("CC","DE") OR GPP_TRSFVENTE="X")' ;
{$ENDIF}

//*** BlocagLanees des natures de pièces autorisées en affaires
if (ctxAffaire in V_PGI.PGIContexte) and (stArgument='ACHAT') then
   begin
   st_plus:=st_Plus +AfPlusNatureAchat;
   if Not(ctxGCAFF in V_PGI.PGIContexte) then begin
      SetControlVisible('TGP_DEPOT',false);
      SetControlVisible('GP_DEPOT',false);
      SetControlVisible('TGP_DOMAINE',false);
      SetControlVisible('GP_DOMAINE',false);
      end;
   end;
// **** Fin affaire ***

THValComboBox (Ecran.FindComponent('GP_NATUREPIECEG')).Plus := St_Plus ;
THValComboBox (Ecran.FindComponent('GP_DEPOT')).ItemIndex := 0;
THValComboBox (Ecran.FindComponent('GP_ETABLISSEMENT')).ItemIndex := 0;
THValComboBox (Ecran.FindComponent('RUPT1')).ItemIndex := 0;

// critére de regroupement affaire caché si pas de contexte affaire PA
if Not(ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) then
   begin
   for i := 1 to 6 do begin SetControlProperty('RUPT'+IntToStr(i),'Plus','AND CO_CODE<>"AFF"'); end;
   end;

if stArgument='VENTE' then
   begin
   SetControlText('GP_NATUREPIECEG','CC');
   SetControlProperty('GP_TIERS','DataType','GCTIERSCLI') ;
   end else
if stArgument='ACHAT' then
   begin
   if not(ctxScot in V_PGI.PGIContexte) then SetControlText('GP_NATUREPIECEG','CF');
   SetControlProperty('GP_TIERS','DataType','GCTIERSFOURN') ;
   SetControlProperty('GP_TIERS_','DataType','GCTIERSFOURN') ;
   { mcd 17/12/01 aussi gérer en entête piece ...// critère affaire non visible sur les pièces d'achats car gérée à la ligne...
   SetControlVisible('GP_AFFAIRE1',False); SetControlVisible('GP_AFFAIRE2',False); SetControlVisible('GP_AFFAIRE3',False); SetControlVisible('TGP_AFFAIRE',False);
   SetControlVisible('GP_AFFAIRE1_',False); SetControlVisible('GP_AFFAIRE2_',False); SetControlVisible('GP_AFFAIRE3_',False); SetControlVisible('TGP_AFFAIRE_',False);
   SetControlVisible('BSELECTAFF1',False); SetControlVisible('BSELECTAFF2',False); }
   end;
if Not(ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) then
   begin
   SetControlVisible ('TGP_AFFAIRE', False);
   SetControlVisible ('GP_AFFAIRE1', False);
   SetControlVisible ('GP_AFFAIRE2', False);
   SetControlVisible ('GP_AFFAIRE3', False);
   SetControlVisible ('GP_AVENANT', False);
   SetControlVisible ('TGP_AFFAIRE_', False);
   SetControlVisible ('GP_AFFAIRE1_', False);
   SetControlVisible ('GP_AFFAIRE2_', False);
   SetControlVisible ('GP_AFFAIRE3_', False);
   SetControlVisible ('GP_AVENANT_', False);
   end;
if not VH_GC.GCMultiDepots then
   begin
   SetControlVisible('TGP_DEPOT',false);
   SetControlVisible('GP_DEPOT',false);
   SetControlProperty('RUPT1','Plus','AND CO_CODE<>"DEP"');
   end;
{$ENDIF} //FIN DU IFDEF BTP

  CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;
  //
  //
  CC:=THValComboBox(GetControl('GP_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

//uniquement en line
{*
   SetControlVisible ('GP_REPRESENTANT', False);
   SetControlVisible ('GP_REPRESENTANT_', False);
   SetControlVisible ('TGP_REPRESENTANT', False);
   SetControlVisible ('TGP_REPRESENTANT_', False);
   SetControlVisible ('GP_ETABLISSEMENT', False);
   SetControlVisible ('TGP_ETABLISSEMENT', False);
   TTabSHeet(GetCOntrol('PLIBRETIERS')).TabVisible := false;
   SetControlText('TGP_AFFAIRE', 'Du chantier');
   SetControlText('TGP_AFFAIRE_', 'Au');
   DefiniRuptureS1;
*}

end;

procedure TOF_PtfPiece.ControleChamp(Champ : String;Valeur : String);
var St_Plus : string ;
begin

  if champ='COTRAITANCE' then
  begin
    Setcontroltext('XX_WHERE', Getcontroltext('XX_WHERE')+' AND AFF_MANDATAIRE IN ("X", "-")');
    ecran.caption := 'Portefeuille pièces Cotraitant';
  end
  else if champ='SOUSTRAITANCE' then
  begin
    if copy(ecran.name,1,13) = 'BTPTFPIECECOT' then
    begin
      Setcontroltext('XX_WHERE', Getcontroltext('XX_WHERE')+'(SELECT COUNT(BPI_TIERSFOU) FROM PIECEINTERV WHERE BPI_NATUREPIECEG=GP_NATUREPIECEG AND BPI_SOUCHE=GP_SOUCHE AND BPI_NUMERO=GP_NUMERO) > 0');
    end else
    begin
      Setcontroltext('XX_WHERE', Getcontroltext('XX_WHERE')+' AND SSTRAITE > 0');
    end;
    ecran.caption := 'Portefeuille pièces Sous-traitant';
  end
  else if Champ='ACHAT' then
  Begin
    St_Plus:=           ' AND GPP_VENTEACHAT="ACH"';
    st_plus:= st_Plus + ' AND ((GPP_NATUREPIECEG="DEF") or (GPP_NATUREPIECEG="CF") or (GPP_NATUREPIECEG="BLF")';
    st_Plus:= St_PLUS + '  OR (GPP_NATUREPIECEG="AF") or (GPP_NATUREPIECEG="AFS") or (GPP_NATUREPIECEG="FF"))';
    st_Plus:= St_PLUS + '  OR (GPP_NATUREPIECEG="BFA") '; //FV1 - 29/11/2016 - FS#2251 - DELABOUDINIERE - Dans édition portefeuille ajouter les pièces RETOURS FOURNISSEURS
    THValComboBox (Ecran.FindComponent('GP_NATUREPIECEG')).Plus := St_Plus ;
    //
    GestionFS1635('', False);
    //
    SetControlProperty('GP_TIERS','DataType','GCTIERSFOURN') ;
    SetControlProperty('GP_TIERS_','DataType','GCTIERSFOURN') ;
    SetControlProperty('GP_NATUREPIECEG','Vide',True);
  End
  else if Champ = 'DEVIS' then
  begin
    st_plus:=st_Plus + ' AND (GPP_NATUREPIECEG="DBT")';
    THValComboBox (Ecran.FindComponent('GP_NATUREPIECEG')).Plus := St_Plus ;
    SetControlEnabled ('GP_NATUREPIECEG',false);
    //FV1 : 25/06/2015 - FS#1635 - SERVAIS : en portefeuille devis, pb sur sélection
    SetControlText    ('TETATAFFAIRE',    'Code Etat');
    SetControlVisible ('AFF_ETATAFFAIRE', True);
    SetControlVisible ('AFF_GENERAUTO',   False);
  end
  else if Champ = 'FACTURE' then
  begin
    SetControlProperty('GP_NATUREPIECEG','Plus',' AND (GPP_NATUREPIECEG="ABT" OR GPP_NATUREPIECEG="FBT")');
    SetControlProperty('GP_NATUREPIECEG','Vide',True);
    //FV1 : 25/06/2015 - FS#1635 - SERVAIS : en portefeuille devis, pb sur sélection
    SetControlText    ('TETATAFFAIRE',    'Type Fact.');
    SetControlVisible ('AFF_ETATAFFAIRE', False);
    SetControlVisible ('AFF_GENERAUTO',   True);
  end
  else if Champ = 'COMPTOIR' then
  begin
    st_plus:=st_Plus + ' AND ((GPP_NATUREPIECEG="FBC") or (GPP_NATUREPIECEG="ABC"))';
    THValComboBox (Ecran.FindComponent('GP_NATUREPIECEG')).Plus := St_Plus ;
    SetControlProperty('GP_NATUREPIECEG','Vide',True);
    GestionFS1635('', False);
  end;

  //FV1 : 25/06/2015 - FS#1635 - SERVAIS : en portefeuille devis, pb sur sélection
  {else if Champ = 'STATUT' then
  begin
    SetControlText('TETATAFFAIRE','Type Fact.');
    SetControlVisible('AFF_ETATAFFAIRE',False);
    SetControlVisible('AFF_GENERAUTO',True);
  end;}

  //
  if Champ='STATUT' then
  Begin
    if Valeur = 'APP' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'W')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'W');
      SetControlText('XX_WHERE', Getcontroltext('XX_WHERE')+' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("","W")');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Appel');
      SetControlText('TGP_AFFAIRE', 'Code Appel');
    end
    else if Valeur = 'INT' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'I')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'I');
      SetControlText('XX_WHERE', Getcontroltext('XX_WHERE')+' AND ((SUBSTRING(GP_AFFAIRE,1,1)="I") OR (GP_GENERAUTO="CON"))');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Contrat');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Contrat');
      SetControlText('TGP_AFFAIRE', 'Code Contrat');
    End
    else if Valeur = 'GRP' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('XX_WHERE', Getcontroltext('XX_WHERE')+' AND AFF_AFFAIRE0 IN ("W","A")')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('XX_WHERE', Getcontroltext('XX_WHERE')+' AND SUBSTRING(GP_AFFAIRE,1,1)IN ("W","A")');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Affaire');
      SetControlText('TGP_AFFAIRE', 'Code Affaire');
    end
    else if Valeur = 'AFF' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'A')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'A');
      SetControlText('XX_WHERE', Getcontroltext('XX_WHERE')+' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("A", "")');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Chantier');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Chantier');
      SetControlText('TGP_AFFAIRE', 'Code Chantier');
    end
    else if Valeur = 'PRO' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'P')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'P');
      SetControlText('XX_WHERE', Getcontroltext('XX_WHERE')+' AND SUBSTRING(GP_AFFAIRE,1,1)="P"');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel d''offre');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Appel d''offre');
      SetControlText('TGP_AFFAIRE', 'Code Appel d''offre');
    end
    else
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', '')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', '');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Affaire');
      SetControlText('TGP_AFFAIRE', 'Code Affaire');
    end;
  end;

end;

Procedure TOF_PtfPiece.GestionFS1635(Lib : String; Affiche : Boolean);
begin

  SetControlText    ('TETATAFFAIRE',    Lib);
  SetControlVisible ('AFF_ETATAFFAIRE', Affiche);
  SetControlVisible ('AFF_GENERAUTO',   Affiche);

end;

procedure TOF_PtfPiece.OnExitArticle(Sender : Tobject);
begin

  if (CodeArticle.Text <> '') And (CodeArticle1.Text = '') then
    CodeArticle1.Text := CodeArticle.Text;

end;


Procedure TOF_PtfPiece.NomsChampsAffaire ( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ;
BEGIN
Aff:=THEdit(GetControl('GP_AFFAIRE'))   ; Aff0:=Nil ;
Aff1:=THEdit(GetControl('GP_AFFAIRE1')) ; Aff2:=THEdit(GetControl('GP_AFFAIRE2')) ;
Aff3:=THEdit(GetControl('GP_AFFAIRE3')) ; Aff4:=THEdit(GetControl('GP_AVENANT'))  ;
Aff_:=THEdit(GetControl('GP_AFFAIRE_'))   ; Aff0_:=Nil ;
Aff1_:=THEdit(GetControl('GP_AFFAIRE1_')) ; Aff2_:=THEdit(GetControl('GP_AFFAIRE2_')) ;
Aff3_:=THEdit(GetControl('GP_AFFAIRE3_')) ; Aff4_:=THEdit(GetControl('GP_AVENANT_'))  ;
END ;

procedure TOF_PtfPiece.OnLoad;
begin
    Inherited;
{$IFNDEF BTP}
if GetControlText('GP_NATUREPIECEG') = '' then
    BEGIN
    LastError:=1;
    LastErrorMsg:=TexteMessage[LastError] ;
    exit;
    END;
{$ENDIF}


// JT, eQualité 10782
//FV1 : 14/03/2017 - FS#2375 - LEHODEY TP - Problème édition portefeuille avec article contenant le symbole _
  if CodeArticle <> nil then
  begin
    SetControlText('GL_CODEARTICLE', CodeArticle.text);
    if CodeArticle.text <> '' then
    Begin
      If CodeArticle1.text = '' then CodeArticle1.text := trim(CodeArticle.text);
      //CodeArticle1.text := format('%-18.18s%16.16s',[CodeArticle1.Text,StringOfChar('Z',16)])
    end
    else
      codeArticle1.text := '';
    SetControlText('GL_CODEARTICLE_', copy(trim(CodeArticle1.Text)+StringOfChar('Z',18),1,18));
  end;

  SetControlText('XX_ORDERBY', 'GP_SOUCHE,GP_NUMERO');

end;

procedure TOF_PtfPiece.OnUpdate;
begin
    Inherited;
{$IFNDEF BTP}
if GetControlText('GP_NATUREPIECEG') = '' then
    BEGIN
    END;
{$ENDIF}
end;

/////////////////////////////////////////////////////////////////////////////
Procedure TOF_PtfPiece_AffectGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value, St_Text, St : string;
    Indice, i_ind : integer;
BEGIN
F := TForm (Longint (Parms[0]));
(*
{$IFDEF BTP}
if (F.Name <> 'BTPTFPIECE') then exit;
{$ELSE}
if (F.Name <> 'GCPTFPIECE') then exit;
{$ENDIF}
*)
Indice := Integer (Parms[1]);
St_Plus := string (THValComboBox (F.FindComponent('RUPT1')).Plus);
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
St_Text := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text);
For i_ind := 1 to 6 do
    BEGIN
    if i_ind = Indice then continue;
    St := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value);
    If St <> '' then St_Plus := St_Plus + ' AND CO_CODE <>"' + St + '"';
    END;
{$IFDEF CCS3}
PlusS3:='AND CO_CODE<>"AFF" AND (CO_CODE<"LT5" OR CO_CODE>"LTA")' ;
St_Plus:=St_Plus+PlusS3 ;
{$ENDIF}
if (ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) then St_Plus:=St_Plus +'AND CO_CODE<>"DEP" '  ;
//uniquement en line
//St_Plus:=' AND CO_CODE NOT IN ("COM","ETA","LT1","LT2","LT3","LT4","LT5","DEP","LIV")';

THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Plus := St_Plus;
if St_Value<>'' then THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value := St_Value
                else THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).ItemIndex := 0;
END;


Procedure TOF_PtfPiece_ChangeGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value : string;
    Indice, i_ind : integer;
    Entete, Detail : boolean ;
BEGIN
F := TForm (Longint (Parms[0]));
(*
{$IFDEF BTP}
if (F.Name <> 'BTPTFPIECE') then exit;
{$ELSE}
if (F.Name <> 'GCPTFPIECE') then exit;
{$ENDIF}
*)
Indice := Integer (Parms[1]);
St_Plus := '';
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
if St_Value = '' then
    BEGIN
    THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(Indice))).Text := '';
    THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text := '';
    TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Checked := False;
    TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Enabled := False;
    For i_ind := Indice + 1 to 6 do
        BEGIN
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value := '';
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Enabled := False;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Color := clBtnFace;
        THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(i_ind))).Text := '';
        THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(i_ind))).Text := '';
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Checked := False;
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Enabled := False;
        END;
    END else
    BEGIN
    if Indice < 6 then
        BEGIN
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).Color := clWindow;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).ItemIndex := 0;
        END;
    Entete := TCheckBox (F.FindComponent('ENTETEPIECE')).Checked;
    Detail := TCheckBox (F.FindComponent('DETAILLIGNE')).Checked;
    if (Entete) OR (Detail) then
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Enabled := True;
//uniquement en line
{*
    TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Checked := True;
    TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Enabled := false;
*}
    THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(Indice))).Text :=
            RechDom ('GCGROUPPTFPIECE', St_Value, True);
    THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text :=
            string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text);
    END;
END;

Procedure TOF_PtfPiece_AffectSautPage (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Value : string;
    i_ind : integer;
    Entete, Detail : boolean ;
BEGIN
F := TForm (Longint (Parms[0]));
(*
{$IFDEF BTP}
if (F.Name <> 'BTPTFPIECE') then exit;
{$ELSE}
if (F.Name <> 'GCPTFPIECE') then exit;
{$ENDIF}
*)
Entete := TCheckBox (F.FindComponent('ENTETEPIECE')).Checked;
Detail := TCheckBox (F.FindComponent('DETAILLIGNE')).Checked;
if (not Entete) AND (not Detail) then
    BEGIN
    For i_ind := 1 to 6 do
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Enabled := False;
    END else
    BEGIN
    For i_ind := 1 to 6 do
        BEGIN
        St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value);
        if St_Value <> '' then
            TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Enabled := True;
        END;
    END;
END;

Procedure TOF_PtfPiece_ChangeSautPage (parms:array of variant; nb: integer ) ;
var F: TForm;
    i_ind, Indice : integer;
BEGIN
F := TForm (Longint (Parms[0]));
(*
{$IFDEF BTP}
if (F.Name <> 'BTPTFPIECE') then exit;
{$ELSE}
if (F.Name <> 'GCPTFPIECE') then exit;
{$ENDIF}
*)
Indice := Integer (Parms[1]);
if TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Checked then
    BEGIN
    For i_ind := 1 to Indice - 1 do
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Checked := True;
    END else
    BEGIN
    For i_ind := Indice + 1 to 6 do
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Checked := False;
    END;
END;

/////////////////////////////////////////////////////////////////////////////

procedure InitTOFPtfPiece ();
begin
RegisterAglProc('PtfPiece_ChangeGroup', True , 1, TOF_PtfPiece_ChangeGroup);
RegisterAglProc('PtfPiece_AffectGroup', True , 1, TOF_PtfPiece_AffectGroup);
RegisterAglProc('PtfPiece_AffectSautPage', True , 0, TOF_PtfPiece_AffectSautPage);
RegisterAglProc('PtfPiece_ChangeSautPage', True , 0, TOF_PtfPiece_ChangeSautPage);
end;

//uniquement en line
{*
procedure TOF_PtfPiece.DefiniRuptureS1;
begin
  THValComboBox(GetCOntrol('RUPT1')).Plus := ' AND CO_CODE NOT IN ("COM","ETA","LT1","LT2","LT3","LT4","LT5","DEP","LIV")';
  THValComboBox(GetCOntrol('RUPT2')).Plus := ' AND CO_CODE NOT IN ("COM","ETA","LT1","LT2","LT3","LT4","LT5","DEP","LIV")';
  THValComboBox(GetCOntrol('RUPT3')).Plus := ' AND CO_CODE NOT IN ("COM","ETA","LT1","LT2","LT3","LT4","LT5","DEP","LIV")';
  THValComboBox(GetCOntrol('RUPT4')).Plus := ' AND CO_CODE NOT IN ("COM","ETA","LT1","LT2","LT3","LT4","LT5","DEP","LIV")';
  THValComboBox(GetCOntrol('RUPT5')).Plus := ' AND CO_CODE NOT IN ("COM","ETA","LT1","LT2","LT3","LT4","LT5","DEP","LIV")';
  THValComboBox(GetCOntrol('RUPT6')).Plus := ' AND CO_CODE NOT IN ("COM","ETA","LT1","LT2","LT3","LT4","LT5","DEP","LIV")';
end;
*}

Initialization
registerclasses([TOF_PtfPiece]) ;
InitTOFPtfPiece();
end.
