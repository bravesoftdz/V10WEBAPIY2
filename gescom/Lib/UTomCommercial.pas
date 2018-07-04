unit UTomCommercial;

interface

uses {$IFDEF VER150} variants,{$ENDIF}  StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
      MaineAGL,eFiche,eFichList,
{$ELSE}
      db,HDB,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FichList, Fiche,FE_main,
{$ENDIF}
      forms,sysutils,
      ComCtrls,utilPGI,
      HCtrls,HEnt1,HMsgBox,UTOM, Dialogs,
      Menus, M3FP, EntGC,
      Lookup,
      UTOB, AGLInit ,Ent1, HDimension, AglInitGc,
{$IFDEF NOMADESERVER}
{$IFNDEF EAGLCLIENT}
      uToxClasses,
{$ENDIF}
{$ENDIF}
      UtilGC, ParamSoc
{$IFNDEF CCS3}
      ,UtilConfid  //JT eQualité 10964
{$ENDIF}
       ;

Type
     TOM_Commercial = Class (TOM)
     private
       TypeCommercial : string;
       ChangeType     : Boolean;
       //FV1 : 29/01/2018 - FS#2882 - CLOSSUR - Donner la possibilité de fermer un commercial
       ChkFerme       : TcheckBox;
       DateSupp       : THEdit;
       Action         : Tactionfiche;

{$IFDEF NOMADESERVER}
{$IFNDEF EAGLCLIENT}
       LesSites : TCollectionSites ;
{$ENDIF}
{$ENDIF}
       Procedure AfSpecifApporteur;
       procedure majcontact;
       procedure PositionneEtabUser (NomChamp : string);
       procedure FermeOnClick(Sender: Tobject);
     public
       procedure OnArgument (stArgument : String ) ; override ;
       procedure OnChangeField (F : TField)  ; override ;
       procedure OnNewRecord  ; override ;
       procedure OnUpdateRecord ; override ;
       procedure OnAfterUpdateRecord; override ;
       procedure OnLoadRecord ; override ;
       procedure OnDeleteRecord  ; override ;
       procedure OnClose ; override ;
       procedure MessErrCommun(Nomchamp:String; TypeMessErr : Integer);
       procedure VerifCommission(Valcomi : Double);
       procedure GestionSiteRepresentant ;
     END ;

     TOM_Commission = Class (TOM)
     private
       function RechercheNomCommercial(Commer :string; Var Nom :string):boolean;
       procedure MessErrCommun(Nomchamp:String; TypeMessErr : Integer);
     public
       procedure OnChangeField (F : TField)  ; override ;
       procedure OnNewRecord  ; override ;
       procedure OnUpdateRecord ; override ;
       procedure OnLoadRecord ; override ;
       procedure OnDeleteRecord  ; override ;
       Procedure RechercheArticle;
       Procedure VerifArticle;
       procedure CacheGrillDimension;
       Function  TraiterArticle (CodeArticle : string; var TOBArt : TOB) : Boolean;
       Function  ChoisirDimension (St_CodeArt : string; var TOBArt : tob) : Boolean;
       Function  TrouverArticleSQL (CodeArticle: string; var ToBArt : TOB) : T_RechArt ;
       Procedure AffiGrillCodedimension(TOBArt : TOB);
     END ;

const
// libellés des messages
TexteMessage: array[1..27] of string 	= (
          {1}  'Le pourcentage de commissionnement ne peut être négatif'
          {2} ,'Le libellé doit être renseigné'
          {3} ,'La méthode de calcul de la commission doit être renseignée'
          {4} ,'Code article inconnu'
          {5} ,'Le type de commercial est inconnu'
          {6} ,'Le commercial doit être renseigné'
          {7} ,'La nature de pièce déclenchant le calcul de commissionnement est obligatoire'
          {8} ,'La famille de niveau 1 est obligatoire'
          {9} ,'La famille de niveau 3 ne peut être renseignée sans spécifier également la famille de niveau 2'
         {10} ,'La mise à jour n''a pu s''effectuée correctement. Problème d''attribution de clé'
         {11} ,'Cette fiche existe déjà, vous devez la modifier'
         {12} ,'Le pourcentage de commissionnement ne doit pas excéder 100%'
         {13} ,'La suppression est impossible car le commercial est mentionné dans certains documents'
         {14} ,'La suppression est impossible car l''apporteur est utilisé dans certains documents'
         {15} ,''
         {16} ,''
         {17} ,''
         {18} ,''
         {19} ,''
         // Mode fiche vendeur
         {20} ,'Le code doit être renseigné'
         {21} ,'La suppression est impossible car le vendeur est mentionné dans certains documents'
         {22} ,'Le nom doit être renseigné'
         // Saisie préfixe pour création de tiers pour nomade
         {23} ,'Le préfixe du tiers ne doit pas excéder 3 caractères'
         {24} ,'Ce préfixe est déjà utilisé pour un code d''établissement'
         {25} ,'Ce préfixe est déjà utilisé pour un autre représentant'
         {26} ,'Ce préfixe est déjà utilisé dans les paramètres de la société'
         {27} ,'Ce préfixe est déjà utilisé dans les codes tiers'
              );

var CoordDim  : Array of integer ;
    MemoArticle : string ;


implementation
uses BtpUtil,HrichOle;
{ TOM_Commercial }

procedure TOM_Commercial.OnArgument (stArgument : String ) ;
Var Critere, Champ, Valeur : string;
    X, iInd : integer;
    Ctrl : TControl ;
    CC : THValComboBox;
begin
Inherited;
{$IFDEF NOMADESERVER}
{$IFNDEF EAGLCLIENT}
  LesSites := TCollectionSites.Create( TCollectionSite, True ) ;
{$ENDIF}
{$ENDIF}
AppliqueFontDefaut (THRichEditOle(GetControl('GCL_BLOCNOTE')));
ChangeType := True;
TypeCommercial := '';
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    X:=pos(':',Critere);
    if (X <> 0) then
        BEGIN
        Champ :=copy(Critere,1,X-1);
        Valeur:=Copy (Critere,X+1,length(Critere)-X);
        END
    else Champ := Critere;
    if Champ = 'NOCHANGE_TYPECOMMERCIAL' then ChangeType := False else
    if Champ = 'GCL_TYPECOMMERCIAL'      then Typecommercial := Valeur;

    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;
if Not(ChangeType) then SetControlEnabled ('GCL_TYPECOMMERCIAL',False);

//modif du combo GCL_NATUREPIECEG pour calcul de com sur Facture+Avoirs
Ctrl := GetControl('GCL_NATUREPIECEG') ;
if Ctrl <> nil then
  begin
  {$IFDEF EAGLCLIENT}
  Iind:=THValComboBox(Ctrl).Values.IndexOf('FAC');
  if Iind<0 then Iind:=0;
  THValComboBox(Ctrl).Items.Insert(Iind,TraduireMemoire ( 'Facture + Avoir clients' ) );
  THValComboBox(Ctrl).Values.Insert(Iind,'ZZ1');
  {$ELSE}
  Iind:=THDBValComboBox(Ctrl).Values.IndexOf('FAC');
  if Iind<0 then Iind:=0;
  THDBValComboBox(Ctrl).Items.Insert(Iind,TraduireMemoire ( 'Facture + Avoir clients' ) );
  THDBValComboBox(Ctrl).Values.Insert(Iind,'ZZ1');
  {$ENDIF}
  end;
MakeZoomOLE(ecran.Handle) ;

if not(ctxaffaire in V_PGI.PGIContexte ) then   // Paramétrage des libellés des tables libres
  begin
  GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GCL_VALLIBRE', 3, '');
  GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GCL_DATELIBRE', 3, '');
  end;

  Datesupp := THEdit(GetControl('GCL_DATESUPP'));

  If Assigned(GetControl('GCL_FERME')) then
  begin
    ChkFerme := TCheckBox(GetControl('GCL_FERME'));
    ChkFerme.OnClick := FermeOnClick;
  end;
  
PositionneEtabUser('GCL_ETABLISSEMENT') ;
{$IFNDEF CCS3}
AppliquerConfidentialite(Ecran,'GCL'); //JT eQualité 10964
{$ENDIF}
SetActiveTabSheet('PGeneral');
end;

Procedure TOM_Commercial.FermeOnClick(Sender : Tobject);
begin

  if Not (DS.State in [DSInsert, DsEdit]) then DS.Edit;

  If Chkferme.checked then DateSupp.Text := DateToStr(Now) else DateSupp.Text := DateToStr(idate2099);

end;

procedure TOM_Commercial.OnChangeField(F: TField);
var nomChamp : string ;
begin
  NomChamp := F.FieldName ;
  if ( NomChamp = 'GCL_COMMERCIAL' ) then GestionSiteRepresentant ;
end ;

procedure TOM_Commercial.OnNewRecord;
var CC : THVAlComboBox;
begin
Inherited;
SetControlEnabled('GCL_COMMERCIAL',True);
SetField('GCL_DATESUPP', iDate2099);

if ( (ctxaffaire in V_PGI.PGIContexte ) or (ctxgcaff in V_PGI.PGIContexte ))
    and (not(changetype))   then
    SetField('GCL_TYPECOMMERCIAL', TypeCommercial);  // Type commercial forcé à apporteur

if ctxMode in V_PGI.PGIContexte then
    begin
    SetField('GCL_TYPECOMMERCIAL', 'VEN');  // Type commercial forcé à vendeur
    SetField('GCL_NATUREPIECEG', 'FAC');  // Nature pièce forcé à facture client
    SetField('GCL_ETABLISSEMENT', VH^.EtablisDefaut);  // Initialisé par l'établisst par défaut
    end;

    ChkFerme.Checked := False;
    DateSupp.Text    := DateToStr(Idate2099);

PositionneEtabUser ('GCL_ETABLISSEMENT');

end;

procedure TOM_Commercial.OnUpdateRecord;
var err : integer;
    PrefixeTiers, CodeRep : string;
begin
Inherited;
If GetField('GCL_TYPECOMMERCIAL')='' Then
   BEGIN
   MessErrCommun('GCL_TYPECOMMERCIAL', 5);
   Exit;
   END;
SetField('GCL_COMMERCIAL',Trim(GetControlText('GCL_COMMERCIAL')));
If GetField('GCL_COMMERCIAL')='' Then
   BEGIN
   if ctxMode in V_PGI.PGIContexte then err:= 20 else err:=6;
   MessErrCommun('GCL_COMMERCIAL', err);
   Exit;
   END;
If GetField('GCL_LIBELLE')='' Then
   BEGIN
   if ctxMode in V_PGI.PGIContexte then err:= 22 else err:=2;
   MessErrCommun('GCL_LIBELLE', err);
   Exit;
   END;
//valeur non négative
VerifCommission(GetField('GCL_COMMISSION'));
If GetField('GCL_TYPECOM')='' Then
   BEGIN
   MessErrCommun('GCL_TYPECOM',3);
   Exit;
   END;
If (not (ctxMode in V_PGI.PGIContexte)) and (GetField('GCL_NATUREPIECEG')='') Then
   BEGIN
   MessErrCommun('GCL_NATUREPIECEG', 7);
   Exit;
   END;
//Si le préfixe du code tiers est saisi
if (GetField('GCL_PREFIXETIERS')<>'') then
   begin
   PrefixeTiers := GetField('GCL_PREFIXETIERS');
   CodeRep      := GetField('GCL_COMMERCIAL');
   //Test si le préfixe du code tiers > à 3 alphas
   if Length(PrefixeTiers)>3 then
      begin
      MessErrCommun('GCL_PREFIXTIERS', 23);
      Exit;
   end;
   //Test si déjà saisi dans Etablissement
   if ExisteSql('select ET_ETABLISSEMENT from ETABLISS where ET_ETABLISSEMENT like "'+PrefixeTiers+'%"') then
      begin
      MessErrCommun('GCL_PREFIXETIERS', 24);
      Exit;
      end;
   //Test si déjà saisi pour un autre représentant
   if ExisteSql('select GCL_PREFIXETIERS from COMMERCIAL where GCL_COMMERCIAL<>"'+CodeRep+'" and GCL_PREFIXETIERS like "'+PrefixeTiers+'%"') then
      begin
      MessErrCommun('GCL_PREFIXETIERS', 25);
      Exit;
      end;
   //Test si déjà saisi pour la société
   if GetParamSoc('SO_GCPREFIXETIERS') = PrefixeTiers then
      begin
      MessErrCommun('GCL_PREFIXETIERS', 26);
      Exit;
      end;
   //Test si déjà saisi pour un code tiers
   if ExisteSql('select T_TIERS from TIERS where T_TIERS like "'+PrefixeTiers+'%"') then
      begin
      MessErrCommun('GCL_PREFIXETIERS', 27);
      Exit;
      end;
   end;

   If ChkFerme.Checked then
    SetField('GCL_FERME', 'X')
   else
    SetField('GCL_FERME', '-');

   SetField('GCL_DATESUPP', StrToDate(DateSupp.Text));

end;

procedure TOM_Commercial.OnAfterUpdateRecord;
begin
inherited ;
majcontact;
end;

procedure TOM_Commercial.majcontact;
{var TOB_CON:TOB;}
begin
inherited ;
if ctxAffaire in V_PGI.PGIContexte then exit;
{        // Attendre pour création contact de l'établissement
         // Manque liaison Commercial/contact JCF
TOB_CON:=TOB.Create('CONTACT', NIL, -1);  // Création de la TOB des contacts
TOB_CON.PutValue('C_TYPECONTACT','GCL') ;
TOB_CON.PutValue('C_AUXILIAIRE',GetField('GCL_COMMERCIAL')) ;
TOB_CON.PutValue('C_NUMEROCONTACT','1') ;
TOB_CON.PutValue('C_NATUREAUXI','GCL') ;
TOB_CON.PutValue('C_NOM',GetField('GCL_LIBELLE')) ;
TOB_CON.PutValue('C_PRENOM',GetField('GCL_PRENOM')) ;
TOB_CON.PutValue('C_PRINCIPAL','X') ;
TOB_CON.InsertOrUpdateDB (FALSE);
TOB_CON.free;
}
end;

procedure TOM_Commercial.OnDeleteRecord;
var err : integer;
begin
if (GetField('GPL_TYPECOMMERCIAL') = 'APP') then
    BEGIN
    if (ExisteSQL('SELECT GP_APPORTEUR FROM PIECE WHERE GP_APPORTEUR="'
                    +GetField('GCL_COMMERCIAL')+'"')) or
       (ExisteSQL('SELECT GL_APPORTEUR FROM LIGNE WHERE GL_APPORTEUR="'
                    +GetField('GCL_COMMERCIAL')+'"')) then
       BEGIN
       MessErrCommun('GCL_COMMISSION',14); Exit;
       END;
    if ctxAffaire in V_PGI.PGIContexte then
       BEGIN
       if (ExisteSQL('SELECT AFT_APPORTEUR FROM AFFAIRE WHERE AFT_APPORTEUR="'
                    +GetField('GCL_COMMERCIAL')+'"')) then
         BEGIN
         MessErrCommun('GCL_COMMISSION',15); Exit;
         END;
      END;
    END
else
    BEGIN
    if (ExisteSQL('SELECT GP_REPRESENTANT FROM PIECE WHERE GP_REPRESENTANT="'
                    +GetField('GCL_COMMERCIAL')+'"')) or
       (ExisteSQL('SELECT GL_REPRESENTANT FROM LIGNE WHERE GL_REPRESENTANT="'
                    +GetField('GCL_COMMERCIAL')+'"')) then
       BEGIN
       if ctxMode in V_PGI.PGIContexte then err:= 21 else err:=13;
       MessErrCommun('GCL_COMMISSION', err);
       Exit;
       END;
    END;
end;

procedure TOM_Commercial.OnLoadRecord;
begin
Inherited ;
  GestionSiteRepresentant ;

  if GetField('GCL_COMMERCIAL')<>'' then SetControlEnabled('GCL_COMMERCIAL',False) ;

  if (ctxMode in V_PGI.PGIContexte) then SetControlProperty('GCL_ETABLISSEMENT','Plus','ET_SURSITE="X"');

  if ( (ctxaffaire in V_PGI.PGIContexte ) or (ctxgcaff in V_PGI.PGIContexte ))
    and (not(changetype))   then  AfSpecifApporteur;

  if GetField('GCL_FERME') = 'X' then
    ChkFerme.Checked := True
  else
    ChkFerme.Checked := False;

  DateSupp.text := DateToStr(GetField('GCL_DATESUPP'));

end;

procedure TOM_Commercial.OnClose ;
begin
{$IFDEF NOMADESERVER}
{$IFNDEF EAGLCLIENT}
  if LesSites <> nil then LesSites.free ;
{$ENDIF}
{$ENDIF}
end ;

procedure TOM_Commercial.MessErrCommun(NomChamp:String; TypeMessErr : Integer);
begin
SetFocusControl(NomChamp);
LastError := TypeMessErr;
LastErrorMsg := TexteMessage[LastError];
end;

procedure TOM_Commercial.VerifCommission(ValComi : Double);
begin
If ValComi < 0 Then
   BEGIN
   MessErrCommun('GCL_COMMISSION',1);
   Exit;
   END;
If ValComi > 100 then
   BEGIN
   MessErrCommun('GCL_COMMISSION',12);
   Exit;
   END;
SetField('GCL_COMMISSION',ValComi);
end;

procedure TOM_Commercial.GestionSiteRepresentant ;
var
  Itinerant : TGroupBox ;
{$IFDEF NOMADESERVER}
{$IFNDEF EAGLCLIENT}
  LeSite : TCollectionSite ;
  Commercial : String ;
{$ENDIF}
{$ENDIF}
begin
  // Test s'il faut afficher ou non le GroupBox concernant le site représentant
  Itinerant := TGroupBox( GetControl( 'POP' ) ) ;
  if Itinerant <> nil then
  begin
{$IFDEF NOMADESERVER}
{$IFNDEF EAGLCLIENT}
    LeSite := nil ;
    if LesSites <> nil then
    begin
      Commercial := GetField( 'GCL_COMMERCIAL' ) ;
      LeSite := LesSites.FindVariableValue( '$REP', Commercial, '' ) ;
      if ( LeSite <> nil ) and ( Commercial <> '' ) then
      begin
        SetControlText( 'SITECODE', LeSite.SSI_CODESITE ) ;
        SetControlText( 'SITELIB', LeSite.SSI_LIBELLE ) ;
      end ;
    end ;
    Itinerant.visible := ( LeSite <> nil ) ;
{$ELSE}
    Itinerant.visible := False ;
{$ENDIF}
{$ELSE}
    Itinerant.visible := False ;
{$ENDIF}
  end ;
end ;

procedure TOM_Commercial.AFSpecifApporteur;
BEGIN
if GetField('GCL_TYPECOMMERCIAL')='APP' then
    BEGIN
    SetControlVisible ('GCL_ZONECOM',False); SetControlVisible ('TGCL_ZONECOM',False);
    if ctxScot in V_PGI.PGIContexte then SetControlvisible ('BINFO',False);
    Ecran.Caption := 'Apporteur :'; UpdateCaption(Ecran);
    END;
END;

procedure TOM_Commercial.PositionneEtabUser(NomChamp: string);
Var Etab : String ;
    Forcer : Boolean ;
    CC :THDbValComboBox;
    LastEnabled : boolean;
begin
CC := THDBValCOmboBox (GetCOntrol(NOMChamp));
if CC = nil then exit;
LastEnabled := CC.Enabled ;
if Not VH^.EtablisCpta then Exit ;
if Not CC.Visible then Exit ;
Etab:=VH^.ProfilUserC[prEtablissement].Etablissement ; if Etab='' then Exit ;
Forcer:=VH^.ProfilUserC[prEtablissement].ForceEtab ;
if CC.Values.IndexOf(Etab)<0 then Exit ;
if ((Not CC.Enabled) and (CC.Value<>Etab)) then CC.enabled := true;
SetField(NomChamp,Etab);
if (Forcer) then CC.Enabled:=False ;
end;

{ TOM_Commission }

procedure TOM_Commission.OnChangeField(F: TField);
var RefFamille : THValComboBox;
begin
Inherited;
If (F.FieldName='GCM_FAMILLENIV1') and (DS.State in [dsInsert]) then
    BEGIN
    RefFAmille := THValComboBox (GetControl('GCM_FAMILLENIV1'));
    SetField('GCM_LIBELLE',RefFamille.Text) ;
    END;
if Ecran.Name = 'GCCOMISARTICLE' then
  begin
  if (F.FieldName = 'GCM_CODEARTICLE') then VerifArticle;
  end;
end;

procedure TOM_Commission.OnDeleteRecord;
begin
Inherited;
//Affecter des composants à vide,non visible pr Tlabel
THEdit(GetControl('NomCommercial')).Text := '';
If (Ecran.Name = 'GCCOMISARTICLE') Then
   BEGIN
   //THEdit(GetControl('NOMARTICLE')).Text := '';
   CacheGrillDimension;
   END;
end;

procedure TOM_Commission.OnLoadRecord;
Var ExisteNom : Boolean;
    Nom, RefUnique : string;
    NomComer : THEdit ;
    TOBArt : TOB ;
    QQ : TQuery ;
begin
Inherited;
NomComer := THEdit(GetControl('NOMCOMMERCIAL'));
//Récupérer le nom commercial
If (GetField('GCM_COMMERCIAL')<> '') Then
    BEGIN
    ExisteNom := RechercheNomcommercial(Getfield('GCM_COMMERCIAL'),Nom);
    If Not ExisteNom Then Exit
    Else If ((Uppercase(Ecran.Name)='GCCOMISFAMILLE') or (Uppercase(Ecran.Name)='GCCOMISARTICLE'))
        and (NomComer<>nil) then
        NomComer.Text := Nom;
    END;
If (Ecran.name = 'GCCOMISARTICLE') then
    BEGIN
    //NomArticle  := THEdit(GetControl('NOMARTICLE'));
    RefUnique := GetField('GCM_ARTICLE') ;
    If (RefUnique <>'') then
        BEGIN
        QQ := OpenSQL('Select * from ARTICLE Where GA_ARTICLE="' +
                      RefUnique + '"',True) ;
        if Not QQ.EOF then
            BEGIN
            TOBArt := TOB.Create ('ARTICLE', nil, -1) ;
            TOBArt.SelectDB ('', QQ);
            AffiGrillCodedimension (TOBArt);
           // NomArticle.Text := TOBArt.GetValue ('GA_LIBELLE') ;
            TOBArt.Free ;
            END;
        Ferme (QQ) ;
        END;
    END;
end;

procedure TOM_Commission.OnNewRecord;
begin
Inherited;
If (Ecran.Name = 'GCCOMISARTICLE') Then
    BEGIN
    //THEdit(GetControl('NOMARTICLE')).Text := '';
    MemoArticle := '';
    CacheGrillDimension;
    END;
end;

procedure TOM_Commission.OnUpdateRecord;
var QQ        :TQuery;
    NumeroCom :String;
    Numero    :Integer;
begin
Inherited;
If GetField('GCM_LIBELLE')='' Then
    BEGIN
    MessErrCommun('GCM_LIBELLE',2);
    Exit;
    END;
If GetField('GCM_TYPECOM')='' Then
    BEGIN
    MessErrCommun('GCM_TYPECOM', 3);
    Exit;
    END;
If GetField('GCM_NATUREPIECEG')='' Then
    BEGIN
    MessErrCommun('GCM_NATUREPIECEG', 7);
    Exit;
    END;
If (Ecran.Name ='GCCOMISFAMILLE') then
    BEGIN
    If GetField('GCM_FAMILLENIV1')='' Then
        BEGIN
        MessErrCommun('GCM_FAMILLENIV1', 8);
        Exit;
        END;
    If (GetField('GCM_FAMILLENIV3')<>'') And
       (GetField('GCM_FAMILLENIV1')<>'') And
       (GetField('GCM_FAMILLENIV2')='')  Then
        BEGIN
        MessErrCommun('GCM_FAMILLENIV2', 9);
        Exit;
        END;
    END;
//Test si valeur négative ou > à 100%
if GetField('GCM_COMMISSION') < 0 then
    BEGIN
    MessErrCommun('GCM_COMMISSION',1);
    Exit;
    END;
If GetField('GCM_COMMISSION') > 100 then
   BEGIN
   MessErrCommun('GCM_COMMISSION',12);
   Exit;
   END;
//Existence le code article
If (Ecran.Name ='GCCOMISARTICLE') then
    BEGIN
    If (GetField('GCM_ARTICLE') = '') Then
        BEGIN
        MessErrCommun('GCM_CODEARTICLE', 4);
        Exit;
        END;
    END;
//Incrémenter le champ Numerocom
if (DS.State in [dsInsert]) then
    BEGIN
    QQ := OpenSQL ('SELECT MAX(GCM_NUMEROCOM) FROM COMMISSION',True);
    if Not QQ.EOF then
        BEGIN
        NumeroCom := QQ.Fields[0].AsString;
        if NumeroCom = '' then Numero := 1 else Numero := StrToInt(NumeroCom)+1;
        Setfield('GCM_NUMEROCOM', Numero);
        Ferme(QQ);
        END else
        BEGIN
        Ferme(QQ);
        MessErrCommun('GCM_TYPECOM', 10);
        Exit;
        END;
    if (Ecran.Name ='GCCOMISFAMILLE') then
        BEGIN
        if existeSQL('SELECT GCM_NUMEROCOM FROM COMMISSION WHERE ' +
                     'GCM_COMMERCIAL="' + string(GetField('GCM_COMMERCIAL')) +
                     '" AND GCM_FAMILLENIV1="' + string(GetField('GCM_FAMILLENIV1')) +
                     '" AND GCM_FAMILLENIV2="' + string(GetField('GCM_FAMILLENIV2')) +
                     '" AND GCM_FAMILLENIV3="' + string(GetField('GCM_FAMILLENIV3')) +
                     '" AND GCM_ARTICLE=""') then
            BEGIN
            MessErrCommun('GCM_FAMILLENIV1', 11);
            exit;
            END;
        END else if (Ecran.Name ='GCCOMISARTICLE') then
        BEGIN
        if existeSQL('SELECT GCM_NUMEROCOM FROM COMMISSION WHERE ' +
                     'GCM_COMMERCIAL="' + string(GetField('GCM_COMMERCIAL')) +
                     '" AND GCM_ARTICLE="' + string(GetField('GCM_ARTICLE')) +
                     '" AND GCM_FAMILLENIV1="" AND GCM_FAMILLENIV2=""' +
                     ' AND GCM_FAMILLENIV3=""') then
            BEGIN
            MessErrCommun('GCM_ARTICLE', 11);
            exit;
            END;
        END;
    END;
end;


procedure TOM_Commission.MessErrCommun(NomChamp:String; TypeMessErr : Integer);
begin
SetFocusControl(NomChamp);
LastError := TypeMessErr;
LastErrorMsg := TexteMessage[LastError];
end;

function TOM_Commission.RechercheNomCommercial(Commer: string;
  var Nom: string): boolean;
var Q : TQuery;
begin
Result := True;
Q := OpenSQL ('SELECT GCL_LIBELLE FROM COMMERCIAL WHERE GCL_COMMERCIAL="'+ Commer+'"',True);
If Not Q.Eof then Nom := Q.Fields[0].AsString else Result := False;
Ferme(Q);
end;

Procedure TOM_Commission.RechercheArticle;
var TOBArt : TOB;
    G_CodeArticle, G_Article : THCritMaskEdit;
BEGIN
G_CodeArticle := THCritMaskEdit (GetControl ('GCM_CODEARTICLE'));
G_Article:= THCritMaskEdit.Create (Application);
G_Article.Text := G_CodeArticle.Text;
G_Article.visible:=False;
G_Article.Top := G_CodeArticle.Top;
G_Article.Left := G_CodeArticle.Left;
DispatchRecherche (G_Article, 1, '','GA_CODEARTICLE=' + Trim (Copy (G_CodeArticle.Text, 1, 18)), '');
if G_Article.Text <> '' then
    BEGIN
    TOBArt := TOB.Create ('ARTICLE', nil, -1) ;
    TOBArt.SelectDB ('"' + G_Article.Text + '"', Nil) ;
    G_CodeArticle.Text := TOBArt.GetValue ('GA_CODEARTICLE') ;
    if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
    SetField('GCM_CODEARTICLE',TOBArt.GetValue('GA_CODEARTICLE')) ;
    SetField('GCM_ARTICLE',TOBArt.GetValue('GA_ARTICLE')) ;
    if TraiterArticle (G_CodeArticle.Text, TOBArt) then
        BEGIN
        SetField('GCM_CODEARTICLE',TOBArt.GetValue('GA_CODEARTICLE')) ;
        SetField('GCM_ARTICLE',TOBArt.GetValue('GA_ARTICLE')) ;

        AffiGrillCodedimension(TOBArt);
        MemoArticle := G_CodeArticle.Text ;
        END;
    SetField('GCM_LIBELLE',TOBArt.GetValue('GA_LIBELLE')) ;
    TOBArt.Free ;
    END;
G_Article.Destroy ;
END;

Procedure TOM_Commission.VerifArticle;
var CodeArticle : string;
    TOBArt      : TOB;
BEGIN
CodeArticle := VarAsType(GetField ('GCM_CODEARTICLE'),VarString);
if  CodeArticle='' then
    begin
    MemoArticle:='' ;
    SetField('GCM_LIBELLE','') ;
    SetField('GCM_ARTICLE','') ;
    CacheGrillDimension ;
    end;
if (CodeArticle = '') or (CodeArticle = MemoArticle) or (Not GetControlVisible('GCM_CODEARTICLE')) then
    BEGIN
    exit;
    END else
    BEGIN
    TOBArt := TOB.Create ('ARTICLE', nil, -1) ;
    if TraiterArticle (CodeArticle, TOBArt) then
        BEGIN
//        SetField('GCM_CODEARTICLE',CodeArticle) ;
        SetField('GCM_ARTICLE',TOBArt.GetValue('GA_ARTICLE')) ;
        AffiGrillCodedimension(TOBArt);
        SetField('GCM_LIBELLE',TOBArt.GetValue('GA_LIBELLE')) ;
        MemoArticle := CodeArticle ;
        END else
        BEGIN
        CodeArticle := MemoArticle ;
        SetField('GCM_CODEARTICLE',CodeArticle) ;
        END;
    TOBArt.Free;
    END;
END;

procedure TOM_Commission.CacheGrillDimension ;
Var i_ind : Integer;
begin
For i_ind :=1 To MaxDimension Do
    BEGIN
    SetControlVisible('GRILLIBEL'+ IntToStr(i_ind), False);
    SetControlCaption('GRILLIBEL'+ IntToStr(i_ind), '');
    SetField('CODEDIM'+ IntToStr(i_ind), '');
    SetControlVisible('CODEDIM'+ IntToStr(i_ind), False);
    END;
end;

Function TOM_Commission.TraiterArticle (CodeArticle : string; var TOBArt : TOB) : Boolean;
var RefUnique : string ;
    RechArt : T_RechArt ;
BEGIN
Result := False;
if CodeArticle = '' then
   Begin
   Result := True;
   exit;
   END;
RechArt := TrouverArticleSQL (CodeArticle, TOBArt) ;
Case RechArt of
     traOk    : BEGIN
                Result := True;
                END;
     traGrille: BEGIN
                RefUnique := TOBArt.GetValue ('GA_ARTICLE');
                If ChoisirDimension (RefUnique, TOBArt) then Result := true;
                END;
     END;
END;

function TOM_Commission.ChoisirDimension (St_CodeArt : string; var TOBArt : TOB) : Boolean;
var QQ   :TQuery;
BEGIN
TheTOB:=TOB.Create('', nil, -1);
AglLanceFiche ('GC','GCSELECTDIM','','', 'GA_ARTICLE='+ St_CodeArt +
  	           ';ACTION=SELECT;CHAMP= ') ;
if TheTOB = Nil then
    BEGIN
    Result := False;
    END else
    BEGIN
    Result := True;
    QQ := OpenSQL('Select * from ARTICLE Where GA_ARTICLE="' +
             TheTOB.Detail[0].GetValue ('GA_ARTICLE') +'" ',True) ;
    if Not QQ.EOF then
        BEGIN
        TOBArt.SelectDB ('', QQ);
        END;
    Ferme(QQ) ;
    TheTOB := Nil;
    END;
TheTOB.Free;
END;

Function TOM_Commission.TrouverArticleSQL (CodeArticle: string; var ToBArt : TOB) : T_RechArt ;
var Q    :TQuery;
    Etat :string ;
BEGIN
Result := traAucun;
if CodeArticle = '' Then exit;
Q := OpenSQL('Select * from ARTICLE Where GA_CODEARTICLE="' +
             CodeArticle + '" AND GA_STATUTART <> "DIM" ',True) ;
if Not Q.EOF then
    BEGIN
    TOBArt.SelectDB('',Q);
    Etat:=TOBArt.GetValue('GA_STATUTART') ;
    if Etat='UNI' then Result:=traOk else
        if Etat = 'GEN' then Result:=traGrille else Result:=traOk ;
    END ;
Ferme(Q) ;
END;

Procedure TOM_Commission.AffiGrillCodedimension(TOBArt : TOB);
Var i_ind, i_Dim : Integer;
    GrilleDim,CodeDim,LibDim : String ;
begin
//Affecter les coord des composants
i_Dim := 0;
For i_ind := 1 To MaxDimension do
    BEGIN
    GrilleDim:=TOBArt.GetValue('GA_GRILLEDIM'+IntTostr(i_ind)) ;
    CodeDim:=TOBArt.GetValue('GA_CODEDIM'+IntTostr(i_ind)) ;
    If GrilleDim<>'' Then
       BEGIN
       SetControlCaption('GRILLIBEL'+ IntTostr(i_ind), Rechdom('GCGRILLEDIM'+IntTostr(i_ind),GrilleDim,False));
       LibDim:=GCGetCodeDim(GrilleDim,CodeDim,i_ind) ;
       if LibDim<>'' then
           BEGIN
           SetControlText('CODEDIM'+ IntToStr(i_ind), LibDim);
           SetControlProperty('CODEDIM'+ IntToStr(i_ind), 'Left',CoordDim[i_Dim]);
           SetControlProperty('GRILLIBEL'+ IntTostr(i_ind), 'Left',CoordDim[i_Dim]);
           Inc(i_Dim) ;
           END else
           BEGIN
           SetControlText('CODEDIM'+ IntToStr(i_ind), '');
           SetControlCaption('GRILLIBEL'+ IntTostr(i_ind), '');
           END;
       END else
       BEGIN
       SetControlText('CODEDIM'+ IntToStr(i_ind), '');
       SetControlCaption('GRILLIBEL'+ IntTostr(i_ind), '');
       END;
       SetControlVisible('CODEDIM'+ IntToStr(i_ind), Not (GetControlText('CODEDIM'+ IntToStr(i_ind)) = '')) ;
       SetControlVisible('GRILLIBEL'+ IntToStr(i_ind), Not (GetControlText('CODEDIM'+ IntToStr(i_ind)) = '')) ;
    END;
end;


//----------------------------------------------------------------
Procedure TOMCommercial_InitForm (Parms : Array of variant; nb: integer) ;
var F     : TForm ;
    i_ind : integer ;
    TLIB : THLabel;
    CHPS : THEdit;
BEGIN
F := TForm (Longint (Parms[0]));
if F.Name <> 'GCCOMISARTICLE' then exit;
SetLength (CoordDim, MaxDimension) ;
For i_ind := Low (CoordDim) to High (CoordDim) do
    BEGIN
    TLIB := THLabel(F.FindComponent('GRILLIBEL'+ IntTostr(i_ind + 1)));
    CHPS := THEdit(F.FindComponent('CODEDIM'+IntTostr(i_ind + 1)));
    CoordDim[i_ind]:= CHPS.Left ;
    TLIB.Caption := '';
    CHPS.Text := '';
    CHPS.Visible := Not (CHPS.Text = '');
    END;
MemoArticle := '' ;
END;

procedure COMRechercheArticle (parms: array of variant; nb: integer);
var  F : TForm ;
     OM : TOM ;
     NomChamps : string;
BEGIN
  F:=TForm(Longint(Parms[0])) ;
  NomChamps:=string (Parms [1]);
  if F.Name <> 'GCCOMISARTICLE' then exit;
  if (F is TFFicheListe) then OM:=TFFicheListe(F).OM else exit;
  if (OM is TOM_Commission)
    then TOM_Commission(OM).RechercheArticle
    else exit;
END;

procedure COMVerifArticle (parms: array of variant; nb: integer);
var  F : TForm ;
     OM : TOM ;
     NomChamps : string;
BEGIN
  F:=TForm(Longint(Parms[0])) ;
  NomChamps:=string (Parms [1]);
  if F.Name <> 'GCCOMISARTICLE' then exit;
  if (F is TFFicheListe) then OM:=TFFicheListe(F).OM else exit;
  if (OM is TOM_Commission)
    then TOM_Commission(OM).VerifArticle
    else exit;
END;

procedure InitTOMCommercial ();
begin
RegisterAglProc( 'COMInitForm', True , 0, TOMCommercial_InitForm);
RegisterAglProc( 'COMRechercheArticle', True , 1, COMRechercheArticle);
RegisterAglProc( 'COMVerifArticle', True , 1, COMVerifArticle);
end;

procedure AGLmajcontact_com( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_Commercial)
    then TOM_Commercial(OM).majcontact
    else exit;
end;


Initialization
registerclasses([TOM_Commercial]);
registerclasses([TOM_Commission]);
InitTOMCommercial;
RegisterAglProc( 'majlescontacts_com', TRUE , 0, AGLmajcontact_com); // Mode
end.


