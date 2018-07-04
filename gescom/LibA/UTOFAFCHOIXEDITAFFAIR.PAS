{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : 14/02/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFCHOIXEDITAFFAIR ()
Mots clefs ... : TOF;AFCHOIXEDITAFFAIR
*****************************************************************}
Unit UTOFAFCHOIXEDITAFFAIR ;

Interface

Uses StdCtrls, Controls, Classes,  forms, sysutils,  ComCtrls,
{$IFDEF EAGLCLIENT}
      Maineagl, UtileAGL,
{$ELSE}
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db, FE_Main,
   EdtrEtat, EdtrDoc,
{$ENDIF}

     HCtrls, HEnt1, HMsgBox, UTOF, UTOFAFBASEMODELEEDIT,uRecupSqlModele, Dicobtp;

Type
  TOF_AFCHOIXEDITAFFAIR = Class (TOF_AFBASEMODELEEDIT)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (stArgument : String ) ; override ;
    procedure RBEditParserClick (Sender:TObject);
    public
    sAffaire:string;
  end ;
 Function AFLanceFiche_ChoixImprAff(Argument:string):variant;


implementation

{$IFNDEF EAGLCLIENT}
uses UtofLanceParser;
{$ENDIF}

procedure TOF_AFCHOIXEDITAFFAIR.OnArgument (stArgument:String);
var
   //st,
   Critere, champ, valeur :string;
   x                      :integer;
begin
  Inherited ;
  sAffaire:='';
  Critere:=(Trim(ReadTokenSt(stArgument)));
  While (Critere <>'') do
    BEGIN
    X:=pos('=',Critere);
    if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;

    if (Champ='AFFAIRE') then
	        begin
            sAffaire := Valeur;
	        end;

    if (Valeur='PRO') then
        begin
        Ecran.Caption := TraduitGA('Edition de la proposition');
        end;

    if (Valeur='AFF') then
        begin
        Ecran.Caption := TraduitGA('Edition de l''affaire');
        end;

    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;

    TRadioButton(GetControl('RBEDITPARSER')).OnClick := RBEditParserClick;
{$IFDEF EAGLCLIENT}
    TRadioButton(GetControl('RBEDITPARSER')).Visible := FALSE;
{$ENDIF}

    updatecaption(Ecran);
end ;

procedure TOF_AFCHOIXEDITAFFAIR.OnUpdate ;
var
   TSL              :TStringList;
   TL               :TList;
   stSQL, sDocEtat  :string;
   i                :integer;
   strNatureLettre  :string;
   ModeleEdition    :string;
begin
     inherited ;

     // Il faut une affaire
     if sAffaire = '' then
        exit;

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

          // Lanceur génération doc parser
//          AGLLanceFiche ('AFF', 'AFLANCEPARSER', '', '', 'AFFAIRE=' + sAffaire + ';MODELE=' + ModeleEdition + ';TITRE=Constitution lettre de mission;NATURE=' + strNatureLettre);
{$IFNDEF EAGLCLIENT}
         AFLanceFiche_LanceParser ('', 'AFFAIRE=' + sAffaire + ';MODELE=' + ModeleEdition + ';TITRE=Constitution lettre de mission;NATURE=' + strNatureLettre);
{$ENDIF}
          exit;
     end;

if  TCheckBox(GetControl('CHKEDITFICHE')).Checked = false then
    begin
    sDocEtat:='L';
    if (ModeleEdit.plus = 'APE') or (ModeleEdit.plus = 'AFE') then
        sDocEtat:='E';

    // $$$jp: initialisation de TL
    TL := nil;
    try
    TL := TList.Create ;
    TSL := TStringList.Create ;
    stSQL:=RecupSQLModele(sDocEtat, ModeleEdit.plus, ModeleEdit.value,'','','',' WHERE AFF_AFFAIRE="'+sAffaire+'" ');
    if sDocEtat='E' then
    Begin
    // gm 02/09/02
    // modif pour ALGOE, car dans un état  specif il n'y a pas de jointure sur table ligne
    // donc l'order by fait planter
    // modif provixsoire je regarde si dans la requete on trouve le mot " LIGNE "
    	if pos(' LIGNE ',stSql) > 0 then
    	stSql:=stSql+' order by GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_ARTICLE'
    End
    else
    	stSQL  := stSQL + ' ORDER BY GL_NUMLIGNE';
    TSL.Add (stSql);
    TL.Add (TSL);
    if (ModeleEdit.plus = 'APE') or (ModeleEdit.plus = 'AFE') then
        LanceEtat(sDocEtat,ModeleEdit.plus, ModeleEdit.Value, true, false, false, Nil, trim(stSql), '', False)
    else
        LanceDocument(sDocEtat, ModeleEdit.plus, ModeleEdit.Value, TL, Nil, true, False);

    finally
           for i:=0 to TL.Count-1 do
               TStringList(TL[i]).Free;
           TL.Free;
    end;

    end;

end ;

procedure TOF_AFCHOIXEDITAFFAIR.RBEditParserClick (Sender:TObject);
var
   strPlus    :string;
begin
     if (TRadioButton (GetControl('RBEDITPARSER')).Checked = TRUE) then
     begin
          THValComboBox (GetControl('MODELEEDITION')).plus := '';
          THValComboBox (GetControl('MODELEEDITION')).DataType := 'AFDOCEXTERNE';
     end;

     strPlus := 'ADE_DOCEXTYPE="WOR" AND ADE_DOCEXETAT="UTI" ';
      // mcd 27/04/04 fait dans tous les cas, même sur base mono if V_PGI_ENV.ModeFonc <> 'MONO' then
     strPlus := strPlus + 'AND (ADE_DOSSIER="$STD" OR ADE_DOSSIER="$DAT" OR ADE_DOSSIER="' + V_PGI.NoDossier + '") ';
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


Function AFLanceFiche_ChoixImprAff(Argument:string):variant;
begin
     result := AGLLanceFiche ('AFF','AFCHOIXEDITAFFAIR','','',Argument);
end;

Initialization
  registerclasses ( [ TOF_AFCHOIXEDITAFFAIR ] ) ;
end.




