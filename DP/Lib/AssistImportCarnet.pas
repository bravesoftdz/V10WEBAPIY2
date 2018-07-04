unit AssistImportCarnet;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, ComCtrls, HSysMenu, hmsgbox, StdCtrls, ExtCtrls, HTB97, Hctrls,
  Mask, UTob, UTom, HEnt1, HPanel;

type
  TFAssistImportCarnet = class(TFAssist)
    PChoixFichier: TTabSheet;
    lblSelection: THLabel;
    lblAvert1: THLabel;
    lblAvert2: THLabel;
    PImportation: TTabSheet;
    lblAvert3: THLabel;
    lblFichier: THLabel;
    LeFichier: THCritMaskEdit;
    HLabel1: THLabel;
    lblAide: THLabel;
    procedure bFinClick(Sender: TObject);
  private
    { Déclarations privées }
    function OutlookToRepert(TobSource, TobDest : TOB): Boolean;
    function GetSeparateurSimple(FichierTxt : String): String;
    procedure CAT_CreerTobFichierTexte (var LaTob : Tob; CheminFichier : String;
     Separateur : String; ListeChamp : String);
  public
    { Déclarations publiques }
  end;

procedure LanceAssistImportCarnet ;


///////////// IMPLEMENTATION //////////////
implementation

uses CAT_FichierTexte;

procedure LanceAssistImportCarnet ;
var F : TFAssistImportCarnet ;
BEGIN
  F := TFAssistImportCarnet.Create(Application) ;
  Try
    F.ShowModal ;
  Finally
    F.Free ;
  end ;
end ;

{$R *.DFM}

procedure TFAssistImportCarnet.bFinClick(Sender: TObject);
var TobOutlook, TobRepert : TOB;
    TomRepert : TOM;
    i : Integer;
    Sep : String;
begin
  inherited;
  // Traitement déjà fait
  if BFin.Caption='&Fermer' then begin Self.Close; exit; end;

  // Positionnement sur la page de fin
  RestorePage ;
  P.ActivePage := PImportation ;
  PChange(nil) ;

  // Verifs
  if LeFichier.Text='' then
    begin
    PGIInfo('Aucun fichier choisi !', Self.Text);
    exit;
    end;
  if Not FileExists(LeFichier.Text) then
    begin
    PGIInfo('Le fichier n''existe pas !', Self.Text);
    exit;
    end;

  // 2 types de séparateur possible avec l'export ascii outlook
  Sep := GetSeparateurSimple(LeFichier.Text);
  if Sep='' then
    begin
    PGIInfo('Le fichier n''a pas pu être identifié comme un export Outlook ou Outlook Express.'
     + ' (absence de séparateur Tabulation, Virgule ou Point-virgule)', Self.Text);
    exit;
    end;

  // Traitement commence
  BFin.Caption := '&Fermer';

  // Lecture fichier texte vers tob, avec suppression de tous les guillemets
  // dans les données et les noms de champs
  // et dédoublonnage des noms de champs apparaissant plusieurs fois
  // (bug export outlook avec 2 champs Titre)
  CAT_CreerTobFichierTexte(TobOutlook, LeFichier.Text, Sep, '');

  // Création tob cible
  TobRepert := TOB.Create('YREPERTPERSO', Nil, -1) ;
  TomRepert := CreateTom('YREPERTPERSO', Nil, True, True);

  for i:=0 to TobOutlook.Detail.Count-1 do
    begin
    // Champs systèmes, valeurs par défaut (initnew)
    TobRepert.InitValeurs;

    // Passer par OnNewRecord de la tom
    TomRepert.InitTob(TobRepert);

    // Copie des champs
    if OutlookToRepert(TobOutlook.Detail[i], TobRepert) then
      begin
      // Passe par le OnUpdate de la tom
      if TomRepert.VerifTOB(TobRepert) then
        begin
        // Création enreg
        TobRepert.InsertDB(Nil);

        // Passe par le OnAfterUpdate de la tom
        TomRepert.AfterVerifTOB(TobRepert);
        end;
      end;
    end;

  TobRepert.Free;
  TomRepert.Free;
  TobOutlook.Free;

  // Import terminé, le mul affichera les nouveaux enreg...
  Self.Close;
end;

function TFAssistImportCarnet.OutlookToRepert(TobSource, TobDest : TOB): Boolean;
// Transfère les valeurs des champs Outlook vers YRepertPerso
// Result à True si il y a au moins un nom ou prénom
var CC, Fax : String;
    Tel : TStringList;
begin
  // Le tronquage à la longueur maxi du destinataire se fait tout seul pour les varchar
  // Les champs ne sont pas forcément tous dans le fichier d'export, donc test FielExists
  // La clé YRP_CODEREP, YRP_USER est alimentée par la tom
  Result := False;

  // ---- Dénomination
  if TobSource.FieldExists('Nom') and (TobSource.GetValue('Nom')<>'') then
    begin Result := True; TobDest.PutValue('YRP_NOM', TobSource.GetValue('Nom') ); end;
  // Outlook Express
  // - on prend "Nom"+A0 qui représente le vrai nom,
  //        car "Nom" sans A0 représente un champ concaténé nom_prénom
  // - sauf si champ prénom n'existe pas, dans ce cas garder "Nom" qui est plus riche !
  if TobSource.FieldExists('Nom'+#160) and (TobSource.GetValue('Nom'+#160)<>'')
  and TobSource.FieldExists('Prénom') then
    begin Result := True; TobDest.PutValue('YRP_NOM', TobSource.GetValue('Nom'+#160) ); end;
  // Rien trouvé
  if Not Result then exit;

  // ---- Prénom
  if TobSource.FieldExists('Prénom') and (TobSource.GetValue('Prénom')<>'') then
    TobDest.PutValue('YRP_PRENOM', TobSource.GetValue('Prénom') );

  // ---- Civilité
  // Rq : il y a 2 champs "Titre" dans l'export Outlook !
  // mais heureusement CAT_CreerTobFichierTexte renomme le 2è en Titre_1 ...
  if TobSource.FieldExists('Titre') and (TobSource.GetValue('Titre')<>'') then
    begin
    CC := UpperCase(TobSource.GetValue('Titre'));
    // Tentative de correspondance Outlook (libre!) => PGI (tablette personnalisable !)
    if (CC='DR.') or (CC='DOCTEUR') then CC := 'DR'
    else if (CC='M.') or (CC='MONSIEUR') or (CC='MR.') then CC := 'MR'
    else if (CC='ME') or (CC='MADEMOISELLE') or (CC='MLLE') or (CC='MLLE.') then CC := 'MLE'
    else if (CC='MM') or (CC='MESSIEURS') or (CC='MMS') then CC := 'MM.'
    else if (CC='MME.') or (CC='MADAME') then CC := 'MME';
    if (CC<>'') and (Length(CC)<4) and
     ExisteSQL('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="CIV" AND CC_CODE="'+CC+'"') then
      TobDest.PutValue('YRP_CIVILITE', CC );
    end;

  // ---- N° de téléphone
  Tel := TStringList.Create;
  // ** Outlook **
  if TobSource.FieldExists('Téléphone (bureau)') and (TobSource.GetValue('Téléphone (bureau)')<>'') then
    Tel.Add(TobSource.GetValue('Téléphone (bureau)'));
  if TobSource.FieldExists('Téléphone (domicile)') and (TobSource.GetValue('Téléphone (domicile)')<>'') then
    Tel.Add(TobSource.GetValue('Téléphone (domicile)'));
  if TobSource.FieldExists('Téléphone (autre)') and (TobSource.GetValue('Téléphone (autre)')<>'') then
    Tel.Add(TobSource.GetValue('Téléphone (autre)'));
  if TobSource.FieldExists('Tél. mobile') and (TobSource.GetValue('Tél. mobile')<>'') then
    Tel.Add(TobSource.GetValue('Tél. mobile'));
  // Autres champs disponibles :
  //   Téléphone 2 (bureau) => peu utilisé ?
  //   Téléphone 2 (domicile) => peu utilisé ?
  //   Téléphone (voiture) => obsolète
  //   Téléphone société => pas représentatif du contact
  //   Téléphone principal => peu utilisé ?
  // ** Outlook Express **
  if TobSource.FieldExists('Téléphone professionnel') and (TobSource.GetValue('Téléphone professionnel')<>'') then
    Tel.Add(TobSource.GetValue('Téléphone professionnel'));
  if TobSource.FieldExists('Téléphone personnel') and (TobSource.GetValue('Téléphone personnel')<>'') then
    Tel.Add(TobSource.GetValue('Téléphone personnel'));
  if TobSource.FieldExists('Téléphone mobile') and (TobSource.GetValue('Téléphone mobile')<>'') then
    Tel.Add(TobSource.GetValue('Téléphone mobile'));
  // ** Insertion **
  if Tel.Count>0 then TobDest.PutValue('YRP_TEL1', Tel[0] );
  if Tel.Count>1 then TobDest.PutValue('YRP_TEL2', Tel[1] );
  Tel.Free;

  // ---- N° de fax
  Fax := '';
  // ** Outlook **
  if TobSource.FieldExists('Télécopie (bureau)') then
    Fax := TobSource.GetValue('Télécopie (bureau)');
  if (Fax='') and TobSource.FieldExists('Télécopie (domicile)') then
    Fax := TobSource.GetValue('Télécopie (domicile)');
  if (Fax='') and TobSource.FieldExists('Télécopie (autre)') then
    Fax := TobSource.GetValue('Télécopie (autre)');
  // ** Outlook express **
  if (Fax='') and TobSource.FieldExists('Télécopie professionnelle') then
    Fax := TobSource.GetValue('Télécopie professionnelle');
  if (Fax='') and TobSource.FieldExists('Télécopie personnelle') then
    Fax := TobSource.GetValue('Télécopie personnelle');
  // ** Insertion **
  if Fax<>'' then TobDest.PutValue('YRP_FAX', Fax );

  // ---- Internet
  if TobSource.FieldExists('Adresse de messagerie') then
    TobDest.PutValue('YRP_EMAIL', TobSource.GetValue('Adresse de messagerie') );
end;

function TFAssistImportCarnet.GetSeparateurSimple(FichierTxt : String): String;
// FichierTxt : chemin d'un export txt outlook existant
// Recherche dans la 1ère ligne le séparateur, avec par ordre de préférence :
// tab / point virgule / virgule / rien !
// Corrige cette première ligne (liste de champs) pour corriger l'export Outlook :
//  il contient 2 champs "Titre" !
var F : TextFile;
    S : String;
begin
  Result := '';
  AssignFile(F, FichierTxt);
  Reset(F);
  Readln(F, S);
  CloseFile(F);

  if Pos(#9, S)>0 then
    Result := #9
  else if Pos(';', S)>0 then
    Result := ';'
  else if Pos(',', S)>0 then
    Result := ',';
end;

//-----------------------------------------------------------------------------
//--- Nom   : CAT_CreerTobFichierTexte
//--- Objet : Crée et intialise la tob à partir d'élément d'un fichier texte
//---         Si ChChamp est vide alors on considère que la description
//---         des champs est la première ligne du fichier Texte.
//---         Supprime les guillemets dans les noms de champ et les données
//-----------------------------------------------------------------------------
procedure TFAssistImportCarnet.CAT_CreerTobFichierTexte (var LaTob : Tob; CheminFichier : String;
  Separateur : String; ListeChamp : String);
var FichierTob                    : TFichierTexte;
    ChaineSource                  : AnsiString;
    TobChamp,TobEnreg,TobResultat : Tob;
    Indice,NumChamp               : Integer;
    ChValeur                      : String;
begin
 if Not CAT_OuvrirFichierTexte (FichierTob,CheminFichier,MODE_LIRE) then
  PgiInfo ('Impossible d''ouvrir le fichier '+CheminFichier)
 else
  begin
   if (ListeChamp='') then
    begin
     CAT_LireFichierTexte (FichierTob,ListeChamp);
     ListeChamp:=StringReplace(ListeChamp,Separateur,'|',[rfReplaceAll]);
     ListeChamp:=StringReplace(ListeChamp,'"','',[rfReplaceAll]);

     //--- Création de la tob contenant les noms de champ
     TobChamp:=TOB.Create('', Nil, -1);
     while (ListeChamp<>'') do
      begin
       TobEnreg:=Tob.Create ('',TobChamp,-1);
       TobEnreg.AddChampSupValeur ('Champ',ReadToKenPipe (ListeChamp,'|'));
      end;

     //--- Recherche et remplace les noms de champ identiques
     for Indice:=0 to TobChamp.Detail.Count-1 do
      begin
       NumChamp:=0;
       ChValeur:=TobChamp.Detail [Indice].GetValue ('Champ');
       TobResultat:=TobChamp.FindFirst(['Champ'],[ChValeur],True);

       while TobResultat<>Nil do
        begin
         if (NumChamp>0) then
          TobResultat.PutValue ('Champ',ChValeur+'_'+IntToStr (NumChamp));

         inc (NumChamp);
         TobResultat:=TobChamp.FindNext(['Champ'],[ChValeur],True);
        end;
      end;

     //--- Construit une chaine de caractère contenant la liste des champs
     ListeChamp:='';
     for Indice:=0 to TobChamp.Detail.Count-1 do
      ListeChamp:=ListeChamp+TobChamp.Detail [Indice].GetValue ('Champ')+'|';

     TobChamp.Free;
    end;

   LaTob:=TOB.Create('',Nil,0);
   while not CAT_FinFichierTexte (FichierTob) do
    begin
     CAT_LireFichierTexte (FichierTob,ChaineSource);
     ChaineSource:=StringReplace(ChaineSource,Separateur,'|',[rfReplaceAll]);
     ChaineSource:=StringReplace(ChaineSource,'"','',[rfReplaceAll]);
     TobEnreg:=TOB.Create('',LaTob,-1);
     TobEnreg.LoadFromSt(ListeChamp,'|',ChaineSource,'|');
    end;

   CAT_FermerFichierTexte (FichierTob);
  end;
end;



end.
