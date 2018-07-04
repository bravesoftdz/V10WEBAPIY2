unit uControlCP;

interface

uses
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     uTob,
     HCtrls,
     HEnt1,
     HMsgBox,
     SysUtils,
     uLibWindows,
     UDossierSelect, // MD 27/04/04, voir DP\Lib
     Dialogs;

// Liste des constantes nécessaires aux contrôles
const
  cNomRapportCpt = 'RapportCptCP.log';
  cNomRapportMvt = 'RapportMvtCP.log';
  cNomRapportLet = 'RapportLetCP.log';
  cNomRapportCor = 'RapportCorCp.log';

  cFI_TableCorrect      = 'CORRECDOSSIER';
  cFI_TableVerif        = 'VERIFDOSSIER';
  cFI_TableEtatsChaines = 'ETATSCHAINES';

  cFI_Libelle           = 'TEMPCEGID';

// Prototypage des contrôles
procedure ControleCpt67( vMaj : Boolean ) ; // Controle et correction des comptes 6 et 7

// Prototypage des fonctions d' enregistrement des journaux
function  AjouteErreurCpt ( vSql : TQuery ; vTypeCompte : string ; vMessage : string ; vBoParle : Boolean = False) : Tob;
procedure AjouteErreurMvt ( vTobEnCours, vTobMere : Tob ; vMessage : string ; vBoParle : Boolean = False);
procedure AjouteErreurLet ( vTobEnCours, vTobMere : Tob ; vMessage : string ; vBoParle : Boolean = False);
procedure AjouteErreurCor ( vMessage1 : string; vMessage2 : string ; VBoParle : Boolean = False );

procedure SauveErreur ( vNomRapport : string ; vTobErreur : Tob);
procedure FreeNil ( vObject : TObject );

implementation

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/03/2002
Modifié le ... :   /  /
Description .. : Controle de G_NATUREGENE pour les comptes 6 et 7
Mots clefs ... :
*****************************************************************}
procedure ControleCpt67( vMaj : Boolean ) ;
var lSql : TQuery;
begin
  lSql := nil;
  try
    lSql := OpenSql('Select * from GENERAUX where ' +
                    '(G_GENERAL like "6%" and G_NATUREGENE <> "CHA") or ' +
                    '(G_GENERAL like "7%" and G_NATUREGENE <> "PRO") order by G_GENERAL', True);
    if not lSql.Eof then
    begin
      if not vMaj then
      begin
        SauveErreur( cNomRapportCpt, AjouteErreurCpt( lSql , 'Général', 'La nature du compte n''est pas renseigné correctement' ));
      end
      else
      begin
        try
          BeginTrans;
          ExecuteSql('Update GENERAUX set G_NATUREGENE="CHA" where (G_GENERAL like "6%" and G_NATUREGENE <> "CHA")');
          ExecuteSql('Update GENERAUX set G_NATUREGENE="PRO" where (G_GENERAL like "7%" and G_NATUREGENE <> "PRO")');
          CommitTrans;
        except
          on e: Exception do
          begin
            RollBack;
            PgiInfo('Erreur lors de la modification de la nature des comptes 6 et 7' + #13 + #10 +
                    'Message :' + E.Message, 'Correction des comptes...');
          end;
        end;
      end;
    end;

  finally
    if Assigned( lSql ) then
      Ferme( lSql );

  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/03/2002
Modifié le ... :   /  /
Description .. : Retourne une Tob compati. avec le journal d'erreur des comptes
Mots clefs ... :
*****************************************************************}
function AjouteErreurCpt ( vSql : TQuery ; vTypeCompte : string ; vMessage : string ; vBoParle : Boolean = False ) : Tob;
var lTobMere : Tob;
    lTobFille : Tob;
begin
  Result := nil;

  if vBoParle then
    PgiError( vMessage, '')
  else
  begin
    lTobMere := Tob.Create('', nil, -1);

    vSql.First;
    while not vSql.Eof do
    begin
      lTobFille := Tob.Create('', lTobMere, -1 );
      lTobFille.AddChampSupValeur('Dossier', VH_Doss.Nodossier);
      lTobFille.AddChampSupValeur('Compte', vTypeCompte);
      lTobFille.AddChampSupValeur('Code', vSql.Findfield('G_GENERAL').AsString);
      lTobFille.AddChampSupValeur('Libellé', vSql.FindField('G_LIBELLE').AsString);
      lTobFille.AddChampSupValeur('Remarques', vMessage );
      vSql.Next;
    end;
    Result := lTobMere;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/03/2002
Modifié le ... :   /  /
Description .. : Retourne une Tob compatible avec le jouirnal d'erreur des mvts
Mots clefs ... :
*****************************************************************}
procedure AjouteErreurMvt ( vTobEnCours, vTobMere : Tob ; vMessage : string ; vBoParle : Boolean = False ) ;
var lTobErreur : Tob;
begin
  if vBoParle then
    PgiError( vMessage,'')
  else
  begin
    lTobErreur := TOB.Create('', vTobMere, -1);
    lTobErreur.AddChampSupValeur('Dossier',       VH_Doss.NoDossier);
    lTobErreur.AddChampSupValeur('Journal',       vTobEnCours.GetValue('E_JOURNAL'));
    lTobErreur.AddChampSupValeur('Ref. interne',  vTobEnCours.GetValue('E_REFINTERNE'));
    lTobErreur.AddChampSupValeur('Pièce / Ligne', IntToStr(vTobEnCours.GetValue('E_NUMEROPIECE')) + '/' + IntToStr(vTobEnCours.GetValue('E_NUMLIGNE')));
    lTobErreur.AddChampSupValeur('Date',          DateToStr(vTobEnCours.GetValue('E_DATECOMPTABLE')));
    lTobErreur.AddChampSupValeur('Remarques',     vMessage);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/03/2002
Modifié le ... :   /  /
Description .. : Retourne une Tob compatible avec le jouirnal d'erreur de Lettrage
Mots clefs ... :
*****************************************************************}
procedure AjouteErreurLet ( vTobEnCours, vTobMere : Tob ; vMessage : string ; vBoParle : Boolean = False);
var lTobErreur : Tob;
begin
  if vBoParle then
    PgiError( vMessage,'')
  else
  begin
    lTobErreur := TOB.Create('', vTobMere, -1);
    lTobErreur.AddChampSupValeur('Dossier',       VH_Doss.NoDossier);
    lTobErreur.AddChampSupValeur('Code',          vTobEnCours.GetValue('E_LETTRAGE'));
    lTobErreur.AddChampSupValeur('Auxiliaire',    vTobEnCours.GetValue('E_AUXILIAIRE'));
    lTobErreur.AddChampSupValeur('Général',       vTobEnCours.GetValue('E_GENERAL'));
    lTobErreur.AddChampSupValeur('Etat lettrage', vTobEnCours.GetValue('E_ETATLETTRAGE'));
    lTobErreur.AddChampSupValeur('Remarques',     vMessage + ',Total débit : ' + FormatFloat('#,##0.00',vTobEnCours.GetValue('E_DEBIT'))
                                                  + ' ; Total crédit : ' + FormatFloat('#,##0.00',vTobEnCours.GetValue('E_CREDIT')));
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/04/2002
Modifié le ... :   /  /
Description .. : Ajoute le message dans le journal d' erreurs de correction
Mots clefs ... :
*****************************************************************}
procedure AjouteErreurCor ( vMessage1 : string; vMessage2 : string ; VBoParle : Boolean = False );
var lTobSave, lTobErreur : Tob;
begin
  if vBoParle then
  begin
    PgiError( vMessage1, vMessage2 );
    Exit;
  end;
  // Chargement des erreurs existantes dans le journal
  lTobSave := TOB.Create('', nil, -1);
  if FileExists(GetWindowsTempPath + '\' + cNomRapportCor) then
    TobLoadFromFile(GetWindowsTempPath + '\' + cNomRapportCor, nil, lTobSave);

  lTobErreur := TOB.Create('', lTobSave, -1);
  lTobErreur.AddChampSupValeur('Dossier'    , VH_Doss.NoDossier );
  lTobErreur.AddChampSupValeur('Traitement' , vMessage1 );
  lTobErreur.AddChampSupValeur('Detail'     , vMessage2 );

  lTobSave.SaveToFile(GetWindowsTempPath + '\' + cNomRapportCor, True, True, True);
  FreeNil( lTobErreur );
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/03/2002
Modifié le ... :   /  /
Description .. : Enregistre les filles de vTobErreur dans le journal
Suite ........ : correspondant au critere vNomRapport
Mots clefs ... :
*****************************************************************}
procedure SauveErreur ( vNomRapport : string ; vTobErreur : Tob);
var i : integer;
    lTobSave : Tob;
begin

  // GC
  lTobSave := Tob.Create('', nil, -1);

  // Chargement des erreurs existantes dans le journal
  if FileExists(GetWindowsTempPath + '\' + vNomRapport) then
    TobLoadFromFile(GetWindowsTempPath + '\' + vNomRapport, nil, lTobSave);

  i := 0;
  while i <> vTobErreur.Detail.Count do
  begin
    vTobErreur.Detail[i].ChangeParent(lTobSave, -1);
  end;

  if lTobSave.Detail.Count > 0 then
    lTobSave.SaveToFile(GetWindowsTempPath + '\' + vNomRapport, False, True, True);

  FreeNil( lTobSave );
  (*
  if vTobErreur.Detail.Count > 0 then
    vTobErreur.SaveToFile(GetWindowsTempPath + '\' + vNomRapport, True, True, True); *)

end;

procedure FreeNil(vObject: TObject);
begin
  if Assigned( vObject ) then
  begin
    vObject.Free;
    vObject := nil;
  end;
end;

end.
