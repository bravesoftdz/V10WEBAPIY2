{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 20/10/2000
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFPROPOSEDITMUL ()
Mots clefs ... : TOF;AFPROPOSEDITMUL
*****************************************************************}
Unit UTOFAFPROPOSEDITMUL ;

Interface
Uses  StdCtrls, Controls, Classes,  forms, sysutils,ComCtrls, HTB97, UtilMulTrt,
{$IFDEF EAGLCLIENT}
      eMul,UtileAGL,MaineAGL,
{$ELSE}
      dbTables, db, Mul,HDB, FE_Main,
{$IFDEF V530}
     EdtEtat,EdtDoc,
{$ELSE}
     EdtREtat, EdtRDoc,
{$ENDIF}
{$ENDIF}
      URecupSQLModele,
      UTOB,HCtrls,ParamSoc, HEnt1, HMsgBox, UTOF, DicoAf, uTOFAFBASEMODELEEDIT ;

Type
  TOF_AFPROPOSEDITMUL = Class (TOF_AFBASEMODELEEDIT)
    procedure OnArgument (stArgument : String ) ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
    procedure BOuvrirOnClick(Sender: TObject);
    procedure ZoomOnClick(Sender: TObject);

    procedure RBEditParserClick (Sender:TObject);

    public
      Affaire     :  THEdit;
      Tiers       :  THEdit;
    END ;

const
// libellés des messages
TexteMessage: array[1..3] of string 	= (
          {1}  'Veuillez sélectionner un modèle d''édition'
          {2} ,'Veuillez sélectionner au-moins une '
          {3} ,' dans la liste'
              );
Procedure AFLanceFiche_MulEditPropos(Range,Argument:string);


implementation

uses PGIEnv
{$IFNDEF EAGLCLIENT}
, UtofLanceParser
{$endif}
;


procedure TOF_AFPROPOSEDITMUL.OnArgument (stArgument : String ) ;
var
   //st,
   Critere,champ, valeur :string;
   x                     :integer;
begin
  Inherited ;
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    X:=pos('=',Critere);
    if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;

    if (Champ='STATUT') then
	        begin
            if (Valeur='PRO') then
                begin
                SetControlVisible ('TAFF_ETATAFFAIRE',True);
                SetControlVisible ('AFF_ETATAFFAIRE',True);
                Ecran.Caption := TraduitGA('Edition des propositions');
                TCheckBox(GetControl('AFF_MODELE')).Caption := TraduitGA('Proposition &modèle');
                TCheckBox(GetControl('AFF_ADMINISTRATIF')).Caption := TraduitGA('Proposition &administrative');
                end;

            if (Valeur='AFF') then
                begin
                Ecran.Caption := TraduitGA('Edition des affaires');
                end;
	        end;

    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;

Affaire := THEdit(GetControl('AFFAIRE'));
Tiers := THEdit(GetControl('TIERS'));
TToolBarButton97(GetControl('BPOPMENU')).OnClick := ZoomOnClick;
TToolBarButton97(GetControl('BOuvrir')).OnClick:= BOuvrirOnClick;
TRadioButton(GetControl('RBEDITPARSER')).OnClick := RBEditParserClick;

{$IFDEF EAGLCLIENT}
TRadioButton(GetControl('RBEDITPARSER')).Visible := FALSE;
{$ENDIF}

updatecaption(Ecran);
end ;

procedure TOF_AFPROPOSEDITMUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
    // Le champ AFFAIRE est obligatoire dans la fiche
    // Il sera renseigné par la fct appelante si accès direct à une affaire
    // Ce code ne doit autrement jamais être renseigné par le prog
Aff:=THEdit(GetControl('AFF_AFFAIRE'));
Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
Aff_:=THEdit(GetControl('AFF_AFFAIRE_'));
Aff0_:=THEdit(GetControl('AFF_AFFAIRE0_'));
Aff1_:=THEdit(GetControl('AFF_AFFAIRE1_'));
Aff2_:=THEdit(GetControl('AFF_AFFAIRE2_'));
Aff3_:=THEdit(GetControl('AFF_AFFAIRE3_'));
Aff4_:=THEdit(GetControl('AFF_AVENANT_'));
Tiers:=THEdit(GetControl('AFF_TIERS'));
Tiers_:=THEdit(GetControl('AFF_TIERS_'));
end;

procedure TOF_AFPROPOSEDITMUL.RBEditParserClick (Sender:TObject);
var
   strPlus    :string;
begin
     if (TRadioButton (GetControl('RBEDITPARSER')).Checked = TRUE) then
     begin
          THValComboBox (GetControl('MODELEEDITION')).plus     := '';
          THValComboBox (GetControl('MODELEEDITION')).DataType := 'AFDOCEXTERNE';
     end;

     strPlus := 'ADE_DOCEXTYPE="WOR" AND ADE_DOCEXETAT="UTI" ';
     if V_PGI_ENV.ModeFonc <> 'MONO' then
         strPlus := strPlus + 'AND (ADE_DOSSIER="$STD" OR ADE_DOSSIER="$DAT" OR ADE_DOSSIER="' + V_PGI_ENV.NoDossier + '") ';
     strPlus := strPlus + 'AND ADE_DOCEXNATURE=';

     if (Statut='PRO') then
     begin
          THValComboBox (GetControl('MODELEEDITION')).plus := strPlus + '"PMI"';
          THValComboBox (GetControl('MODELEEDITION')).ItemIndex := 0;
     end
     else
         if (Statut='AFF') then
         begin
              THValComboBox (GetControl('MODELEEDITION')).plus := strPlus + '"LMI"';
              THValComboBox (GetControl('MODELEEDITION')).ItemIndex := 0;
         end;
end;

procedure TOF_AFPROPOSEDITMUL.ZoomOnClick(Sender: TObject);
begin
     //Affaire.Text := FListe.DataSource.DataSet.FindField('AFF_AFFAIRE').AsString;
     //Tiers.Text := FListe.DataSource.DataSet.FindField('AFF_TIERS').AsString;
end;


procedure TOF_AFPROPOSEDITMUL.BOuvrirOnClick(Sender: TObject);
var   i:integer;
      ModeleEdition, {CodeAffaire,} stSQL, LibAffaire, St,sDocEtat : string;
      TSL:TStringList;
      TL:TList;
      Bap : boolean;
      TobMAff : TOB;
      strNatureLettre  :string;
begin
     // JP: lettre ou proposition mission par parser
     if (TRadioButton (GetControl('RBEDITPARSER')).Checked = TRUE) then
     begin
          ModeleEdition := THValComboBox(GetControl('MODELEEDITION')).Value;
          if ModeleEdition = '' then
          begin
               PGIInfo('Veuillez sélectionner un modèle' + #10 + ' Si aucun n''est disponible, vérifiez que vous avez référencé les standards dans le module "paramètres" (modèles d''édition)', Ecran.Caption);
               exit;
          end;

          // Lettre ou proposition de mission
          if Statut = 'AFF' then
              strNatureLettre := 'LMI'
          else
              strNatureLettre := 'PMI';

          TOBMAff := nil;
          try
             // TOB des affaires sélectionnées (avec champ etablissement obligatoire)
             TOBMAff := Tob.Create('les affaires',NIL,-1);
             TraiteEnregMulTable (TFMUL(Ecran), 'SELECT * FROM AFFAIRE', 'AFF_AFFAIRE', 'AFFAIRE', 'AFF_AFFAIRE', 'AFFAIRE', TOBMaff, True);
//             TraiteEnregMulListe (TFMul(Ecran), 'AFF_AFFAIRE', 'AFFAIRE', TOBMAff, True);

             // Lanceur génération doc parser
//             For i:=0 to TobMAff.Detail.Count-1 do
//               begin
//             AGLLanceFiche ('AFF', 'AFLANCEPARSER', '', '', 'TOBAFFAIRE=' + IntToStr (integer (TOBMAff)) + ';MODELE=' + ModeleEdition + ';TITRE=Constitution lettre de mission;NATURE=' + strNatureLettre);
{$IFNDEF EAGLCLIENT}
               AFLanceFiche_LanceParser ('', 'TOBAFFAIRE=' + IntToStr (integer (TOBMAff)) + ';MODELE=' + ModeleEdition + ';TITRE=Constitution lettre de mission;NATURE=' + strNatureLettre);
{$ENDIF}
  //             end
          finally
                 TOBMAff.Free;
          end;

          exit;
     end;



sDocEtat:='L';
if (THValComboBox(GetControl('MODELEEDITION')).plus = 'APE') or (THValComboBox(GetControl('MODELEEDITION')).plus = 'AFE') then
    sDocEtat:='E';
Bap := TCheckBox(GetControl('BAPERCU')).Checked;
if (Statut<>'AFF') then  LibAffaire:='proposition' else LibAffaire:='affaire';
ModeleEdition := THValComboBox(GetControl('MODELEEDITION')).Value;
if (ModeleEdition='') then
    begin
    PGIInfoAf(textemessage[1], Ecran.Caption);
    ModeleEdit.SetFocus; Exit;
    end;

St:= 'Confirmez-vous l''édition de ces ' + LibAffaire + 's';
If (PGIAskAf(st,Ecran.caption)<> mrYes) then exit;

TOBMAff := nil;
TL      := nil;
try
SourisSablier;
TobMAff := Tob.Create('les affaires',NIL,-1);
TraiteEnregMulListe (TFMul(Ecran), 'AFF_AFFAIRE','AFFAIRE', TobMAff, True);
TL := TList.Create ;

For i:=0 to TobMAff.Detail.Count-1 do
   begin
   if TobMAff.detail[i].GetValue('AFF_AFFAIRE') <> '' then
      begin
      stSQL:=RecupSQLModele(sDocEtat, THValComboBox(GetControl('MODELEEDITION')).plus, ModeleEdition,'','','',' WHERE AFF_AFFAIRE="'+TobMAff.detail[i].GetValue('AFF_AFFAIRE')+'" ');
    // gm 02/09/02
    // modif pour ALGOE, car dans un état  specif il n'y a pas de jointure sur table ligne
    // donc l'order by fait planter
    // modif provixsoire je regarde si dans la requete on trouve le mot " LIGNE "
      if pos(' LIGNE ',stSql) > 0 then
      Begin
      if (THValComboBox(GetControl('MODELEEDITION')).plus = 'APE') or (THValComboBox(GetControl('MODELEEDITION')).plus = 'AFE') then
         stSQL  := stSQL + ' ORDER BY GL_NUMLIGNE'
      else
      	 stSql:=stSql+' order by GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_ARTICLE';
      End;
      TSL := TStringList.Create ;
      TSL.Add (stSql); TL.Add (TSL);
      end;
   end;

if (TL.count<>0) then
    if (THValComboBox(GetControl('MODELEEDITION')).plus = 'APE') or (THValComboBox(GetControl('MODELEEDITION')).plus = 'AFE') then
        for i:=0 to TL.Count-1 do
            begin
            stSql :=TStringList(TL[i])[0];
            LanceEtat(sDocEtat,THValComboBox(GetControl('MODELEEDITION')).plus, ModeleEdition, Bap, false, false, Nil, trim(stSql), '', False);
            end
    else
        LanceDocument(sDocEtat, THValComboBox(GetControl('MODELEEDITION')).plus, ModeleEdition, TL, Nil, Bap, False);
      
finally
       TobMaff.Free;
       for i:=0 to TL.Count-1 do TStringList(TL[i]).Free;
           TL.Free;
       SourisNormale;
end;

end;

Procedure AFLanceFiche_MulEditPropos(Range,Argument:string);
begin
AGLLanceFiche ('AFF','AFPROPOSEDITMUL',Range,'',Argument);
end;

Initialization
registerclasses ([TOF_AFPROPOSEDITMUL]) ;
end.
