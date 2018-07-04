unit RecupS3;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Buttons, Db, DBTables, HCtrls, Hent1, UTOB, Hdimension,
  ImgList, ExtCtrls, Ent1, HTB97, HMsgBox;

type
  TFRecupS3 = class(TForm)
    CB1: TCheckBox;
    SB1: TStatusBar;
    PB1: TProgressBar;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BAide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BValider: TToolbarButton97;
    LDepart: TLabel;
    LFin: TLabel;
    LDepartSuite: TLabel;
    procedure bValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { DéclaRatios privées }
    OkOk : boolean ;
    procedure Traite_Tiers;
    procedure Traite_Adresses;
    procedure Traite_Articles;
    procedure Traite_Nomenclatures;
    procedure Traite_Tarif(TypeTarf : string);
    procedure Traite_Tablettes;
    procedure AffecteTAuxiliaire(Var StAuxiliaire : string);
    Procedure RemplaceLesBadCaractere(Var Lib : String) ;
  public
    { DéclaRatios publiques }
  end;

Procedure Lance_Import ;

var
  FRecupS3: TFRecupS3;
  CodeClient : string;
  CodeFour : string;
  bTraite : boolean;
  TobCC : TOB;
  FicLog : textfile;
  st_log : string;

Const
  LgTiers : integer = 731 ;
  LgProd  : integer = 551 ;
  LgClil  : integer = 197 ;
  LgNmcl  : integer = 104 ;
  LgTarf  : integer = 127 ;
  LgTrfo  : integer = 155 ;
      {$IFDEF TESTS1}
  LgPrm2  : integer = 41 ;
      {$ELSE}
  LgPrm2  : integer = 61 ;
      {$ENDIF}

implementation

{$R *.DFM}

Procedure Lance_Import ;
var
   Fo_Import : TFRecupS3;
Begin
     Fo_Import := TFRecupS3.Create (Application);
     Try
         Fo_Import.ShowModal;
     Finally
         Fo_Import.free;
     End;
end;

procedure TFRecupS3.FormShow(Sender: TObject);
begin
CB1.Checked := False;
CB1.Visible := V_PGI.SAV;
SB1.Panels.Items[0].Width := SB1.ClientWidth - 100;
PB1.Visible := False;
LDepart.Visible := True;
LFin.Visible := False;
bTraite := False;
bValider.OnClick := bValiderClick;
end;

procedure TFRecupS3.bValiderClick(Sender: TObject);
begin
if bTraite then Exit;
LDepart.Visible := False;
LFin.Visible := True;
LFin.Invalidate;
if CB1.Checked then
    begin
    SB1.SimpleText := 'Effacement de la table Tiers';
    ExecuteSQL('Delete from TIERS');
    ExecuteSQL('Delete from TIERSCOMPL');
    Sleep(2000);
    SB1.SimpleText := 'Effacement de la table Adresses';
    ExecuteSQL('Delete from ADRESSES where ADR_TYPEADRESSE="TIE"');
    Sleep(2000);
    SB1.SimpleText := 'Effacement de la table Contact';
    ExecuteSQL('Delete from CONTACT');
    Sleep(2000);
    SB1.SimpleText := 'Effacement de la table RIB';
    ExecuteSQL('Delete from RIB');
    Sleep(2000);
    SB1.SimpleText := 'Effacement de la table Article';
    ExecuteSQL('Delete from ARTICLE');
    Sleep(2000);
    SB1.SimpleText := 'Effacement de la table Nomenclature';
    ExecuteSQL('Delete from NOMENENT');
    ExecuteSQL('Delete from NOMENLIG');
    Sleep(2000);
    SB1.SimpleText := 'Effacement de la table Tarif';
    ExecuteSQL('Delete from TARIF');
    Sleep(2000);
    SB1.SimpleText := 'Effacement des tablettes';
    ExecuteSQL('Delete from COMMERCIAL');
    ExecuteSQL('Delete from MODEPAIE');
    ExecuteSQL('Delete from MODEREGL');
    ExecuteSQL('Delete from CHOIXCOD where CC_TYPE="FN1" or CC_TYPE="FN2" or CC_TYPE="FN3" or ' +
               'CC_TYPE="LGU" or CC_TYPE="TRC" or CC_TYPE="TAR"');
    Sleep(2000);
    end;
PB1.Visible := True;
AssignFile (FicLog, 'GCIMPORTS3.log');
Rewrite (FicLog);
Try
  OkOk:=True;
  Traite_Tablettes;
  Traite_Tiers;
  Traite_Adresses;
  Traite_Articles;
  Traite_Nomenclatures;
  Traite_Tarif('CLI');
  Traite_Tarif('FOU');
  TobCC.Free;
  if OkOk then
    PGIInfo('Le traitement est terminé.','Récupération de dossier Négoce')
    else
    PGIInfo('Le traitement a rencontré une erreur.','Récupération de dossier Négoce');
  bTraite := True;
  CloseFile(FicLog);
except
  TobCC.Free;
  PGIInfo('Le traitement a rencontré une erreur.','Récupération de dossier Négoce');
  CloseFile(FicLog);
end;
end;

procedure TFRecupS3.AffecteTAuxiliaire(Var StAuxiliaire : string);
Var Tampon : String;
    LgCode : Byte ;
begin
   LgCode:=VH^.Cpta[fbAux].Lg ;
   Tampon:=Uppercase(Trim(StAuxiliaire));
   if (Length(Tampon)<>LgCode) then
      begin
      if Length(Tampon)>LgCode then Tampon:=Copy(Tampon,1,LgCode)
      else if Length(Tampon)<LgCode then StAuxiliaire:=BourreLaDonc(Tampon,fbAux);
      end;
end;

procedure TFRecupS3.Traite_Tiers;
var
    stEnreg, stDate, stSigne,stAuxi : string;
    FicSeq : textfile;
    TobOut, TobTrav, TobTrav2, TobRib, TobCon, TobTCompl, TobTemp : TOB;
    Nb_Lus, i_sizefic, Ratio, ind1, iPos : integer;
    dbTrav : double;
    Searchrec : TSearchRec;
    ExistFic : boolean;
    MessAvance : string ;

    function LitEnreg : string;
    var  Ch : char;
         ind1 : integer;
    begin
    Result := '';
    for ind1 := 1 to LgTiers do
        begin
        Read(FicSeq, Ch);
        Result := Result + Ch;
        end;
    end;

begin
if not FileExists('C:\PGI01\DAT\TiersS3.txt') then Exit;
SB1.SimpleText := 'Traitement de la table Tiers';
//  ouverture du fichier de données
{$I-}
AssignFile(FicSeq, 'C:\PGI01\DAT\TiersS3.txt');
Reset (FicSeq);
{$I+}
ExistFic:=(IOResult = 0);
if Not ExistFic then exit;
stEnreg := LitEnreg;
Nb_Lus := 0;
TobOut := TOB.Create('', nil, -1);
TobRib := TOB.Create('', nil, -1);
TobCon := TOB.Create('', nil, -1);
TobTCompl := TOB.Create('', nil, -1);
MessAvance:='Début de traitement Tiers ';
Try
  //  calcul nb enreg du fichier de données
  FindFirst('C:\PGI01\DAT\TiersS3.txt', faAnyFile, SearchRec);
  i_sizefic := Trunc(SearchRec.Size / Length(stEnreg));
  PB1.Position := 0;
  Ratio := Round(i_sizefic / 100);
  if Ratio = 0 then Ratio := 1;
  st_log := '****** Début du traitement TIERS ******';
  writeln (FicLog, st_log);
  flush (FicLog);
  Reset (FicSeq);
  while not EOF(FicSeq) do
      begin
      stEnreg := LitEnreg;
      writeln (FicLog, Copy(stEnreg,1,50));
      flush (FicLog);

      Inc(Nb_Lus);
      if ((Nb_Lus mod Ratio) = 0) or (Ratio = 1) then PB1.StepIt;
      SB1.Panels.Items[1].Text := IntToStr(Nb_Lus) + ' sur ' + IntToStr(i_Sizefic);
      TobTrav := TOB.Create('TIERS', TobOut, -1);
      TobTrav.InitValeurs;
      TobTrav2 := TOB.Create('TIERSCOMPL', TobTCompl, -1);
      TobTrav2.InitValeurs;
  //  mise en place des valeurs importées
      MessAvance:='Affectation du code auxiliaire #13 Enregistrement n° '+IntToStr(Nb_Lus);
      {$IFDEF TESTS1}
      stAuxi:= Trim(Copy(stEnreg, 2, 10));
      {$ELSE}
      stAuxi:=Trim(Copy(stEnreg, 1, 1) + Trim(Copy(stEnreg, 2, 10)));
      {$ENDIF}
      AffecteTAuxiliaire(stAuxi);
      TobTrav.PutValue('T_AUXILIAIRE'     , stAuxi);   // <  0> code client     (clé primaire)
  {    stDate := TobTrav.GetValue('T_AUXILIAIRE');
      if Length(stDate) < VH^.Cpta[fbAux].Lg then
      begin
         for ind1 := Length(stDate) to VH^.Cpta[fbAux].Lg do stDate := stDate+VH^.Cpta[fbAux].Cb ;
         end ;
      TobTrav.PutValue('T_AUXILIAIRE',stDate); }
      if CB1.Checked then
          begin
          TobTrav.InitValeurs;
          TobTrav.PutValue('T_AUXILIAIRE',stAuxi);
          end
          else
          begin
          if ExisteSQL('Select T_AUXILIAIRE from Tiers where T_AUXILIAIRE="' + stAuxi + '"') then
                                            TobTrav.LoadDB;
          end;

      MessAvance:='Affectation de la nature du code auxiliaire #13 Enregistrement n° '+IntToStr(Nb_Lus);
      TobTrav.PutValue('T_NATUREAUXI', 'XXX');
      {$IFDEF TESTS1}
      if (Copy(stEnreg, 2, 1) = '9') or (Copy(stEnreg, 2, 1) = 'C') then
          begin
          CodeClient := Copy(stEnreg, 2, 1);
      {$ELSE}
      if (Copy(stEnreg, 1, 1) = '9') or (Copy(stEnreg, 1, 1) = 'C') then
          begin
          CodeClient := Copy(stEnreg, 1, 1);
      {$ENDIF}
          TobTrav.PutValue('T_NATUREAUXI', 'CLI');
          if TobTrav.GetValue('T_COLLECTIF') = '' then TobTrav.PutValue('T_COLLECTIF', VH^.DefautCli);
          end;
      {$IFDEF TESTS1}
      if (Copy(stEnreg, 2, 1) = '0') or (Copy(stEnreg, 2, 1) = 'F') then
          begin
          CodeFour := Copy(stEnreg, 2, 1);
      {$ELSE}
      if (Copy(stEnreg, 1, 1) = '0') or (Copy(stEnreg, 1, 1) = 'F') then
          begin
          CodeFour := Copy(stEnreg, 1, 1);
      {$ENDIF}
          TobTrav.PutValue('T_NATUREAUXI', 'FOU');
          if TobTrav.GetValue('T_COLLECTIF') = '' then TobTrav.PutValue('T_COLLECTIF', VH^.DefautFou);
          end;
      stDate := TobTrav.GetValue('T_COLLECTIF');
      if Length(stDate) > VH^.Cpta[fbGene].Lg then stDate := Copy(stDate,1,VH^.Cpta[fbGene].Lg) ;
      if Length(stDate) < VH^.Cpta[fbGene].Lg then
      begin
         for ind1 := Length(stDate) to VH^.Cpta[fbGene].Lg do stDate := stDate+VH^.Cpta[fbGene].Cb ;
         end ;
      TobTrav.PutValue('T_COLLECTIF',stDate);

      MessAvance:='Affectation code tiers et adresse #13 Enregistrement n° '+IntToStr(Nb_Lus);
      {$IFDEF TESTS1}
      TobTrav.PutValue('T_TIERS'          , Trim(Copy(stEnreg, 2, 10)));   // <  0> code client     (clé primaire)
      {$ELSE}
      TobTrav.PutValue('T_TIERS'          , Copy(stEnreg, 1, 1) + Trim(Copy(stEnreg, 2, 10)));   // <  0> code client     (clé primaire)
      {$ENDIF}
      TobTrav.PutValue('T_JURIDIQUE'      , Trim(Copy(stEnreg, 12, 3)));          // < 10> forme juridique
      TobTrav.PutValue('T_LIBELLE'        , Trim(Copy(stEnreg, 17, 35)));         // < 15> nom du client
      TobTrav.PutValue('T_ABREGE'         , Trim(Copy(stEnreg, 17, 35)));         // < 15> nom du client
      TobTrav.PutValue('T_ADRESSE1'       , Trim(Copy(stEnreg, 52, 35)));         // < 50> adresse 1
      TobTrav.PutValue('T_ADRESSE2'       , Trim(Copy(stEnreg, 87, 35)));         // < 85> adresse 2
      TobTrav.PutValue('T_CODEPOSTAL'     , Trim(Copy(stEnreg, 122, 5)));          // <120> code postal
      TobTrav.PutValue('T_VILLE'          , Trim(Copy(stEnreg, 127, 35)));         // <125> ville
      TobTrav.PutValue('T_PAYS'           , Trim(Copy(stEnreg, 162, 3)));          // <160> code pays
      TobTrav.PutValue('T_TELEPHONE'      , Trim(Copy(stEnreg, 165, 25)));         // <163> téléphone
      TobTrav.PutValue('T_TELEX'          , Trim(Copy(stEnreg, 190, 25)));         // <188> télex
      TobTrav.PutValue('T_FAX'            , Trim(Copy(stEnreg, 215, 25)));         // <213> télécopie
      TobTrav.PutValue('T_TELEX'          , Trim(Copy(stEnreg, 240, 15)));         // <238> code accès minitel
      TobTrav.PutValue('T_LANGUE'         , Trim(Copy(stEnreg, 277, 3)));          // <275> code langue étrangère
      TobTrav.PutValue('T_APE'            , Trim(Copy(stEnreg, 280, 4)));          // <278> code APE client
      TobTrav.PutValue('T_MODEREGLE'      , Trim(Copy(stEnreg, 284, 3)));          // <282> code règlement
      TobTrav.PutValue('T_FACTUREHT'      , 'X');
      TobTrav.PutValue('T_LETTRABLE'      , 'X');
      TobTrav.PutValue('T_CONFIDENTIEL'   , '-');
      TobTrav.PutValue('T_SOLDEPROGRESSIF', 'X');
      TobTrav.PutValue('T_CREERPAR'       , 'IMP');

      MessAvance:='Affectation mode relevé de facture et risque #13 Enregistrement n° '+IntToStr(Nb_Lus);
      if (Copy(stEnreg, 287, 1) = 'O') then
          begin
          TobTrav.PutValue('T_RELEVEFACTURE', 'X');          // <285> code relevé fact (O/N)
          TobTrav.PutValue('T_FREQRELEVE', '1M');
          TobTrav.PutValue('T_JOURRELEVE', 28);
          end
          else
          begin
          TobTrav.PutValue('T_RELEVEFACTURE', '-');
          TobTrav.PutValue('T_FREQRELEVE', '');
          TobTrav.PutValue('T_JOURRELEVE', 0);
          end;

      if (Copy(stEnreg, 288, 1) = 'R') then          // <286> code situation
          TobTrav.PutValue('T_ETATRISQUE', 'R')
      else if (Copy(stEnreg, 288, 1) = 'O') then
          TobTrav.PutValue('T_ETATRISQUE', 'O')
          else
          TobTrav.PutValue('T_ETATRISQUE', 'V');

      MessAvance:='Affectation régime TVA, TPF et devise #13 Enregistrement n° '+IntToStr(Nb_Lus);
      if (Trim(Copy(stEnreg, 525, 3)) <> '') then
          TobTrav.PutValue('T_REGIMETVA', 'CEE')          // <287> régime taxe (T=taxe;E=exo;S=susp.)
      else if (Copy(stEnreg, 289, 1) = 'E') then
          TobTrav.PutValue('T_REGIMETVA', 'EXP')          // <287> régime taxe (T=taxe;E=exo;S=susp.)
      else if (Copy(stEnreg, 289, 1) = 'S') then
          TobTrav.PutValue('T_REGIMETVA', 'EXO')          // <287> régime taxe (T=taxe;E=exo;S=susp.)
      else
          TobTrav.PutValue('T_REGIMETVA', 'FRA');         // <287> régime taxe (T=taxe;E=exo;S=susp.)

      TobTrav.PutValue('T_SOUMISTPF', '-');
      if (Copy(stEnreg, 289, 1) = 'T') and (Copy(stEnreg, 498, 1) = '1') then
              TobTrav.PutValue('T_SOUMISTPF', 'X');

      TobTrav.PutValue('T_DEVISE', Copy(stEnreg, 290, 3));          // <288> code devise de factuRatio
      if Trim(Copy(stEnreg, 290, 3)) = '' then TobTrav.PutValue('T_DEVISE', 'EUR');

      MessAvance:='Affectation tiers facturé, tarif et représentant #13 Enregistrement n° '+IntToStr(Nb_Lus);
      StAuxi:=Trim(Copy(stEnreg, 293, 10));
      if stAuxi<>'' then
         begin
         AffecteTAuxiliaire(stAuxi);
         TobTrav.PutValue('T_FACTURE', StAuxi);        // <291> code client à facturer
         end else
         TobTrav.PutValue('T_FACTURE', TobTrav.GetValue('T_AUXILIAIRE'));

      TobTrav.PutValue('T_TARIFTIERS', Trim(Copy(stEnreg, 324, 3)));          // <322> code tarif
      TobTrav.PutValue('T_COMMENTAIRE', Trim(Copy(stEnreg, 373, 30)));        // <371> commentaires libre 1
      TobTrav.PutValue('T_REPRESENTANT', Trim(Copy(stEnreg, 454, 3)));          // <452> code représentant 1 client
      TobTrav.PutValue('T_NIF', Trim(Copy(stEnreg, 505, 20))); // <503> Numéro d'identification TVA CEE
      TobTrav.PutValue('T_SIRET', Trim(Copy(stEnreg, 548, 17)));         // <546> numero SIRET (cle secondaire 2)

      MessAvance:='Affectation dates création, suppression, fermeture #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stDate := Copy(stEnreg, 470, 8);         // <468> date de creation
      while Pos('/', stDate) <> 0 do Delete(stDate, Pos('/', stDate), 1);
      if (IsNumeric(stDate)) and (Length(stDate) = 6) and (StrToInt(stDate) <> 0) then
          begin
          try
          TobTrav.PutValue('T_DATEOUVERTURE', Str6ToDate(stDate, 20));
          except
              TobTrav.PutValue('T_DATEOUVERTURE', iDate1900);
          end;
          end
          else
          TobTrav.PutValue('T_DATEOUVERTURE', iDate1900);

      stDate := Copy(stEnreg, 442, 8);          // <440> date suppression
      while Pos('/', stDate) <> 0 do Delete(stDate, Pos('/', stDate), 1);
      if (IsNumeric(stDate)) and (Length(stDate) = 6) and (StrToInt(stDate) <> 0) then
          begin
          try
          TobTrav.PutValue('T_DATEFERMETURE', Str6ToDate(stDate, 20));
          except
              TobTrav.PutValue('T_DATEFERMETURE', idate2099);
          end;
          end
          else
          TobTrav.PutValue('T_DATEFERMETURE', idate2099);

      if TobTrav.GetValue('T_DATEFERMETURE') <> idate2099 then
          TobTrav.PutValue('T_FERME', 'X')
          else
          TobTrav.PutValue('T_FERME', '-');

      MessAvance:='Affectation remise, escompte et plafond d''encours #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stDate := Trim(Copy(stEnreg, 634, 13));          // <631> % remise globale client
      stSigne := Copy(stEnreg, 633, 1);
      iPos:=Pos('.', stDate);
      if iPos>0 then stDate[iPos] := ',';
      dbTrav := 0.0;
      if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
      if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
      TobTrav.PutValue('T_REMISE', dbTrav);

      stDate := Trim(Copy(stEnreg, 690, 13));          // <687> % escompte client
      stSigne := Copy(stEnreg, 689, 1);
      iPos:=Pos('.', stDate);
      if iPos>0 then stDate[iPos] := ',';
      dbTrav := 0.0;
      if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
      if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
      TobTrav.PutValue('T_ESCOMPTE', dbTrav);

      stDate := Trim(Copy(stEnreg, 676, 13));          // <673> en cours maxi
      stSigne := Copy(stEnreg, 675, 1);
      iPos:=Pos('.', stDate);
      if iPos>0 then stDate[iPos] := ',';
      dbTrav := 0.0;
      if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
      if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
      TobTrav.PutValue('T_CREDITACCORDE', dbTrav * 1000);
      TobTrav.PutValue('T_CREDITPLAFOND', dbTrav * 1000);

      MessAvance:='Affectation des tables libres tiers #13 Enregistrement n° '+IntToStr(Nb_Lus);
      TobTrav2.PutValue('YTC_AUXILIAIRE', TobTrav.GetValue('T_AUXILIAIRE'));
      TobTrav2.PutValue('YTC_TIERS', TobTrav.GetValue('T_TIERS'));
      TobTrav2.PutValue('YTC_TABLELIBRETIERS1', Trim(Copy(stEnreg, 315, 3)));          // <313> code stat 1
      TobTrav2.PutValue('YTC_TABLELIBRETIERS2', Trim(Copy(stEnreg, 318, 3)));          // <316> code stat 2
      TobTrav2.PutValue('YTC_TABLELIBRETIERS3', Trim(Copy(stEnreg, 321, 3)));          // <319> code stat 3

      MessAvance:='Affectation du RIB #13 Enregistrement n° '+IntToStr(Nb_Lus);
      if Trim(Copy(stEnreg, 361, 11)) <> '' then
          begin
          TobTemp := TOB.Create('RIB', TobRib, -1);
          TobTemp.InitValeurs;
          TobTemp.PutValue('R_AUXILIAIRE', TobTrav.GetValue('T_AUXILIAIRE'));
          TobTemp.PutValue('R_NUMERORIB', Nb_Lus);
          TobTemp.PutValue('R_DOMICILIATION', Trim(Copy(stEnreg, 327, 24)));         // <325> domiciliation bancaire
          TobTemp.PutValue('R_ETABBQ', Trim(Copy(stEnreg, 352, 5)));          // <350> code établissement
          TobTemp.PutValue('R_GUICHET', Trim(Copy(stEnreg, 357, 5)));          // <355> code guichet
          TobTemp.PutValue('R_NUMEROCOMPTE', Trim(Copy(stEnreg, 362, 11)));         // <360> n° de compte
          TobTemp.PutValue('R_CLERIB', Trim(Copy(stEnreg, 726, 2)));         // <688> clé RIB
          end;
  //
      MessAvance:='Affectation du contact #13 Enregistrement n° '+IntToStr(Nb_Lus);
      if Trim(Copy(stEnreg, 254, 20)) <> '' then
          begin
          TobTemp := TOB.Create('CONTACT', TobCon, -1);
          TobTemp.InitValeurs;
          TobTemp.PutValue('C_TYPECONTACT', 'T');
          TobTemp.PutValue('C_AUXILIAIRE', TobTrav.GetValue('T_AUXILIAIRE'));
          TobTemp.PutValue('C_NUMEROCONTACT', Nb_Lus);
          TobTemp.PutValue('C_NATUREAUXI', TobTrav.GetValue('T_NATUREAUXI'));
          TobTemp.PutValue('C_NOM', Trim(Copy(stEnreg, 254, 20)));         // <253> nom personne à contacter
          end;
  //
      MessAvance:='Mise à jour dans la base #13 Enregistrement n° '+IntToStr(Nb_Lus);
      if ((Nb_Lus mod Ratio) = 0) then
          begin
          TobOut.SetAllModifie(True);
          TobTCompl.SetAllModifie(True);
          TobRib.SetAllModifie(True);
          TobCon.SetAllModifie(True);
          if CB1.Checked then
              begin
              TobOut.InsertDB(nil);
              TobTCompl.InsertDB(nil);
              TobRib.InsertDB(nil);
              TobCon.InsertDB(nil);
              end
              else
              begin
              TobOut.InsertOrUpdateDB;
              TobTCompl.InsertOrUpdateDB;
              TobRib.InsertOrUpdateDB;
              TobCon.InsertOrUpdateDB;
              end;
          TobOut.ClearDetail;
          TobTCompl.ClearDetail;
          TobRib.ClearDetail;
          TobCon.ClearDetail;
          end;
      end;
  MessAvance:='Fin du traitement #13 Enregistrement n° '+IntToStr(Nb_Lus);
  CloseFile(FicSeq);
  TobOut.SetAllModifie(True);
  TobTCompl.SetAllModifie(True);
  TobRib.SetAllModifie(True);
  TobCon.SetAllModifie(True);
  if CB1.Checked then
      begin
      TobOut.InsertDB(nil);
      TobTCompl.InsertDB(nil);
      TobRib.InsertDB(nil);
      TobCon.InsertDB(nil);
      end
      else
      begin
      TobOut.InsertOrUpdateDB;
      TobTCompl.InsertOrUpdateDB;
      TobRib.InsertOrUpdateDB;
      TobCon.InsertOrUpdateDB;
      end;
  TobOut.Free;
  TobTCompl.Free;
  TobRib.Free;
  TobCon.Free;
Except
  CloseFile(FicSeq);
  TobOut.Free; TobTCompl.Free;
  TobRib.Free; TobCon.Free;
  writeln (FicLog, 'Erreur : ' + MessAvance);
  MessAvance:=MessAvance+' #13 '+Copy(stEnreg,1,50);
  PGIInfo('Erreur : ' + MessAvance,'Récupération de dossier Négoce');
  OkOk:=False;
end;
end;

procedure TFRecupS3.Traite_Adresses;
var
    stEnreg, stCode,stSql : string;
    FicSeq : textfile;
    TobOut, TobTrav, TobCode, TobTmp : TOB;
    Nb_Lus, i_sizefic, Ratio, MaxAdr : integer;
    Searchrec : TSearchRec;
    ExistFic : boolean;
    QQ : TQuery ;
    MessAvance : string ;

    function LitEnreg : string;
    var  Ch : char;
         ind1 : integer;
    begin
    Result := '';
    for ind1 := 1 to LgClil do
        begin
        Read(FicSeq, Ch);
        Result := Result + Ch;
        end;
    end;

begin
if not FileExists('C:\PGI01\DAT\ClilS3.txt') then Exit;
SB1.SimpleText := 'Traitement de la table ADRESSES';
//  ouverture du fichier de données
{$I-}
AssignFile(FicSeq, 'C:\PGI01\DAT\ClilS3.txt');
Reset (FicSeq);
{$I+}
ExistFic:=(IOResult = 0);
if Not ExistFic then exit;
stEnreg := LitEnreg;
Nb_Lus := 0;
TobOut := TOB.Create('', nil, -1);
TobCode:= TOB.Create('', nil, -1);
MessAvance:='Début de traitement Adresses ';
Try
  //  calcul nb enreg du fichier de données
  FindFirst('C:\PGI01\DAT\ClilS3.txt', faAnyFile, SearchRec);
  i_sizefic := Trunc(SearchRec.Size / Length(stEnreg));
  PB1.Position := 0;
  Ratio := Round(i_sizefic / 100);
  if Ratio = 0 then Ratio := 1;
  st_log := '****** Début du traitement ADRESSES ******';
  writeln (FicLog, st_log);
  flush (FicLog);
  Reset (FicSeq);
  stSql:='Select Max(adr_numeroadresse) from adresses where adr_typeadresse="TIE"';
  QQ:=OpenSQL(stSql,True) ;
  if QQ.Eof then MaxAdr:=0
            else MaxAdr:=QQ.Fields[0].Asinteger;
  Ferme(QQ) ;
  while not EOF(FicSeq) do
      begin
      stEnreg := LitEnreg;
      writeln (FicLog, Copy(stEnreg,1,50));
      flush (FicLog);

      Inc(Nb_Lus); Inc(MaxAdr);
      if ((Nb_Lus mod Ratio) = 0) or (Ratio = 1) then PB1.StepIt;
      SB1.Panels.Items[1].Text := IntToStr(Nb_Lus) + ' sur ' + IntToStr(i_Sizefic);
      TobTrav := TOB.Create('ADRESSES', TobOut, -1);
      TobTrav.InitValeurs;
  //  mise en place des valeurs importées
      MessAvance:='Affectation référence de l''adresse par le code tiers #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stCode:=Uppercase(Trim(Copy(stEnreg, 1, 1) + Trim(Copy(stEnreg, 2, 10))));
      if TobCode.FindFirst(['REFCODE'],[stCode],False) = Nil then
        begin
        TobTmp:=TOB.Create('Code Tiers', TobCode, -1);
        TobTmp.AddChampSupValeur('REFCODE',stCode,False);
        stSql:='Delete adresses where adr_typeadresse="TIE" and adr_refcode="'+stCode+'"' ;
        ExecuteSql(stSql);
        end;
      TobTrav.PutValue('ADR_REFCODE'      , stCode);   // <  0> code client     (clé primaire)
      TobTrav.PutValue('ADR_TYPEADRESSE'  , 'TIE');
      MessAvance:='Affectation de l''adresse complète #13 Enregistrement n° '+IntToStr(Nb_Lus);
      TobTrav.PutValue('ADR_NUMEROADRESSE', MaxAdr);
      TobTrav.PutValue('ADR_JURIDIQUE'    , Trim(Copy(stEnreg, 16, 3)));          // < 15> forme juridique
      TobTrav.PutValue('ADR_LIBELLE'      , Trim(Copy(stEnreg, 21, 35)));         // < 20> nom du client
      TobTrav.PutValue('ADR_ADRESSE1'     , Trim(Copy(stEnreg, 56, 35)));         // < 55> adresse 1
      TobTrav.PutValue('ADR_ADRESSE2'     , Trim(Copy(stEnreg, 91, 35)));         // < 90> adresse 2
      TobTrav.PutValue('ADR_CODEPOSTAL'   , Trim(Copy(stEnreg, 126, 5)));          // <125> code postal
      TobTrav.PutValue('ADR_VILLE'        , Trim(Copy(stEnreg, 131, 35)));         // <130> ville
      TobTrav.PutValue('ADR_PAYS'         , Trim(Copy(stEnreg, 166, 3)));          // <165> code pays
      TobTrav.PutValue('ADR_TELEPHONE'    , Trim(Copy(stEnreg, 169, 25)));         // <168> téléphone

      MessAvance:='Mise à jour dans la base #13 Enregistrement n° '+IntToStr(Nb_Lus);
      if ((Nb_Lus mod Ratio) = 0) then
          begin
          TobOut.SetAllModifie(True);
          if CB1.Checked then
              TobOut.InsertDB(nil)
              else
              TobOut.InsertOrUpdateDB;
          TobOut.ClearDetail;
          end;
      end;
  MessAvance:='Fin du traitement #13 Enregistrement n° '+IntToStr(Nb_Lus);
  CloseFile(FicSeq);
  TobOut.SetAllModifie(True);
  if CB1.Checked then
      TobOut.InsertDB(nil)
      else
      TobOut.InsertOrUpdateDB;
  TobOut.Free; TobCode.Free;
Except
  CloseFile(FicSeq);
  TobOut.Free;  TobCode.Free;
  writeln (FicLog, 'Erreur : ' + MessAvance);
  MessAvance:=MessAvance+' #13 '+Copy(stEnreg,1,50);
  PGIInfo('Erreur : ' + MessAvance,'Récupération de dossier Négoce');
  OkOk:=False;
end;
end;

Procedure TFRecupS3.RemplaceLesBadCaractere(Var Lib : String) ;
var i : integer ;
BEGIN
i:=0 ; if Length(Lib)<=0 then exit ;
While i<Length(Lib) do
  BEGIN
  Inc(i) ;
  If (Lib[i]>=Chr(34)) and (Lib[i]<=Chr(42)) then Lib[i]:=Chr(45);
  END ;
END ;

procedure TFRecupS3.Traite_Articles;
var
    stEnreg, stDate, stSigne, stTemp, stAuxi : string;
    FicSeq : textfile;
    TobOut, TobTrav,TOBNomen : TOB;
    Nb_Lus, iDim, i_sizefic, Ratio,iPos : integer;
    dbTrav : double;
    Searchrec : TSearchRec;
    ExistFic : boolean;
    MessAvance : string ;

    function LitEnreg : string;
    var  Ch : char;
         ind1 : integer;
    begin
    Result := '';
    for ind1 := 1 to LgProd do
        begin
        Read(FicSeq, Ch);
        Result := Result + Ch;
        end;
    end;

begin
if not FileExists('C:\PGI01\DAT\ProdS3.txt') then Exit;
SB1.SimpleText := 'Traitement de la table Article';
//  ouverture du fichier de données
{$I-}
AssignFile(FicSeq, 'C:\PGI01\DAT\ProdS3.txt');
Reset (FicSeq);
{$I+}
ExistFic:=(IOResult = 0);
if Not ExistFic then exit;
stEnreg := LitEnreg;
Nb_Lus := 0;
TobOut := TOB.Create('', nil, -1);
MessAvance:='Début de traitement Article ';
Try
  //  calcul nb enreg du fichier de données
  FindFirst('C:\PGI01\DAT\ProdS3.txt', faAnyFile, SearchRec);
  i_sizefic := Trunc(SearchRec.Size / Length(stEnreg));
  PB1.Position := 0;
  Ratio := Round(i_sizefic / 100);
  if Ratio = 0 then Ratio := 1;
  st_log := '****** Début du traitement ARTICLE ******';
  writeln (FicLog, st_log);
  flush (FicLog);
  Reset (FicSeq);
  while not EOF(FicSeq) do
      begin
      stEnreg := LitEnreg;
  writeln (FicLog, Copy(stEnreg,1,50));
  flush (FicLog);
      Inc(Nb_Lus);
      if ((Nb_Lus mod Ratio) = 0) or (Ratio = 1) then PB1.StepIt;
      SB1.Panels.Items[1].Text := IntToStr(Nb_Lus) + ' sur ' + IntToStr(i_Sizefic);
      TobTrav := TOB.Create('ARTICLE', TobOut, -1);
      TobTrav.InitValeurs;
  //  mise en place des valeurs importées
      MessAvance:='Affectation type de l''article #13 Enregistrement n° '+IntToStr(Nb_Lus);
      if (Copy(stEnreg, 236, 1) = 'N') then
         TobTrav.PutValue('GA_TYPEARTICLE', 'NOM')
         else
         TobTrav.PutValue('GA_TYPEARTICLE', 'MAR');
      TobTrav.PutValue('GA_STATUTART', 'UNI');
      MessAvance:='Affectation code article et article #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stTemp:=Copy(stEnreg, 1, 15) ;
      RemplaceLesBadCaractere(stTEmp);
      TobTrav.PutValue('GA_CODEARTICLE'   ,Uppercase(stTemp) );        // <  0>  code article   (cle primaire)
  //    stDate := Uppercase(CodeArticleUnique2(TobTrav.GetValue('GA_CODEARTICLE'), ''));
      iDim := 18 + (MaxDimension * 3);
      stDate := format('%-'+IntToStr(iDim)+'.'+IntToStr(iDim)+'sX',[TobTrav.GetValue('GA_CODEARTICLE')]);

      TobTrav.PutValue('GA_ARTICLE', stDate);
      MessAvance:='Affectation du libellé article #13 Enregistrement n° '+IntToStr(Nb_Lus);
      TobTrav.PutValue('GA_LIBELLE'       , Trim(Copy(stEnreg, 16, 70))); // < 15>  désignation 1+2+3
      TobTrav.PutValue('GA_LIBCOMPL'      , Trim(Copy(stEnreg, 86, 20))); // < 20>  désignation 3
      TobTrav.PutValue('GA_TARIFARTICLE'  , Trim(Copy(stEnreg, 106, 3)));         // <105>  famille tarif  (cle secondaire 2)
      if (Copy(stEnreg, 124, 1) = 'O') then
          TobTrav.PutValue('GA_COMMISSIONNABLE', 'X')   // <123>  article commissionnable (O/N)
          else
          TobTrav.PutValue('GA_COMMISSIONNABLE', '-');  // <123>  article commissionnable (O/N)
      TobTrav.PutValue('GA_LIBREART1'     , Trim(Copy(stEnreg, 128, 2)));          // <127>  nature d'article
  //  TobTrav.PutValue('GA_COMPTAARTICLE' , Copy(stEnreg, 127, 3));
      MessAvance:='Affectation familles article #13 Enregistrement n° '+IntToStr(Nb_Lus);
      TobTrav.PutValue('GA_FAMILLENIV1'   , Trim(Copy(stEnreg, 130, 3)));        // <129>  code statistique 1
      TobTrav.PutValue('GA_FAMILLENIV2'   , Trim(Copy(stEnreg, 133, 3)));        // <132>  code statistique 2
      TobTrav.PutValue('GA_FAMILLENIV3'   , Trim(Copy(stEnreg, 136, 3)));        // <135>  code statistique 3
      TobTrav.PutValue('GA_QUALIFMARGE'   , 'PC');
      if (Copy(stEnreg, 139, 1) = 'O') then
          TobTrav.PutValue('GA_TENUESTOCK', 'X')   // <138>  tenu en stock  O/N
          else
          TobTrav.PutValue('GA_TENUESTOCK', '-');  // <138>  tenu en stock  O/N
      TobTrav.PutValue('GA_FAMILLETAXE2'  , Copy(stEnreg, 140, 1));      // <139>  code taxe parafiscale
      TobTrav.PutValue('GA_FAMILLETAXE1'  , Copy(stEnreg, 142, 1));         // <141>  code tva des pays de la CEE
      TobTrav.PutValue('GA_CODEDOUANIER'  , Trim(Copy(stEnreg, 162, 10) + Copy(stEnreg, 290, 5)));    // <161>  Code nomenclature de produit (CEE)

      MessAvance:='Affectation date suppression #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stDate := Copy(stEnreg, 205, 8);         // <204> date de suppression
      while Pos('/', stDate) <> 0 do Delete(stDate, Pos('/', stDate), 1);
      if (IsNumeric(stDate)) and (Length(stDate) = 6) and (StrToInt(stDate) <> 0) then
          begin
          try
          TobTrav.PutValue('GA_DATESUPPRESSION', Str6ToDate(stDate, 20));
          except
              TobTrav.PutValue('GA_DATESUPPRESSION', iDate2099);
          end;
          end
          else
          TobTrav.PutValue('GA_DATESUPPRESSION', iDate2099);
      if TobTrav.GetValue('GA_DATESUPPRESSION') <> iDate2099 then
          TobTrav.PutValue('GA_FERME', 'X')
          else
          TobTrav.PutValue('GA_FERME', '-');

      TobTrav.PutValue('GA_CREERPAR', 'IMP');
      TobTrav.PutValue('GA_REMISEPIED', 'X');
      TobTrav.PutValue('GA_ESCOMPTABLE', 'X');
      TobTrav.PutValue('GA_REMISELIGNE', 'X');

      MessAvance:='Affectation coefficient calcul PA -> PV #13 Enregistrement n° '+IntToStr(Nb_Lus);
      if (Copy(stEnreg, 214, 1) = 'V') then
          begin
          TobTrav.PutValue('GA_CALCPRIXHT', 'PAA');   // <213>  mode calcul PV-PA (V, A ou blanc)
          TobTrav.PutValue('GA_CALCAUTOHT', 'X');
          stDate := Trim(Copy(stEnreg, 370, 13));          // <383>  coef. conversion factuRatio
          stSigne := Copy(stEnreg, 369, 1);
          iPos:=Pos('.', stDate);
          if iPos>0 then stDate[iPos] := ',';
          dbTrav := 0.0;
          if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
          if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
          TobTrav.PutValue('GA_COEFCALCHT', dbTrav);
          end
          else
          begin
          TobTrav.PutValue('GA_CALCPRIXHT', 'AUC');  // <213>  mode calcul PV-PA (V, A ou blanc)
          TobTrav.PutValue('GA_CALCAUTOHT', '-');
          TobTrav.PutValue('GA_COEFCALCHT', 0.0);
          end;
      TobTrav.PutValue('GA_CALCPRIXTTC', 'AUC');
      if Trim(Copy(stEnreg, 1, 15)) <> Trim(Copy(stEnreg, 215, 15)) then
         TobTrav.PutValue('GA_CODEBARRE', Trim(Copy(stEnreg, 215, 15)));        // <214>  code à barres  (cle secondaire 3)
      TobTrav.PutValue('GA_REMPLACEMENT', Trim(Copy(stEnreg, 238, 15)));        // <237>  article de remplacement
      TobTrav.PutValue('GA_QUALIFUNITEVTE', Trim(Copy(stEnreg, 256, 2)));       // <255>  unité de factuRatio UF
      TobTrav.PutValue('GA_QUALIFUNITESTO', Trim(Copy(stEnreg, 258, 2)));       // <257>  unite de stock  US
      MessAvance:='Affectation fournisseur principale #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stAuxi := CodeFour + Trim(Copy(stEnreg, 279, 10));
      AffecteTAuxiliaire(stAuxi);
      TobTrav.PutValue('GA_FOURNPRINC', stAuxi);     // <278>  Fournisseur Principal
      MessAvance:='Affectation date création #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stDate := Copy(stEnreg, 260, 8);         // <259> date de creation
      while Pos('/', stDate) <> 0 do Delete(stDate, Pos('/', stDate), 1);
      if (IsNumeric(stDate)) and (Length(stDate) = 6) and (StrToInt(stDate) <> 0) then
          begin
          try
          TobTrav.PutValue('GA_DATECREATION', Str6ToDate(stDate, 20));
          except
              TobTrav.PutValue('GA_DATECREATION', iDate1900);
          end;
          end
          else
          TobTrav.PutValue('GA_DATECREATION', iDate1900);

      MessAvance:='Affectation prix unitaire de vente #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stDate := Trim(Copy(stEnreg, 440, 13));          // <438>  prix unitaire de vente
      stSigne := Copy(stEnreg, 439, 1);
      iPos:=Pos('.', stDate);
      if iPos>0 then stDate[iPos] := ',';
      dbTrav := 0.0;
      if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
      if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
      TobTrav.PutValue('GA_PVHT', dbTrav);
      
      MessAvance:='Affectation prix d''achat (DPA, PMAP) #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stDate := Trim(Copy(stEnreg, 412, 13));          // <410>  dernier prix d'achat
      stSigne := Copy(stEnreg, 411, 1);
      iPos:=Pos('.', stDate);
      if iPos>0 then stDate[iPos] := ',';
      dbTrav := 0.0;
      if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
      if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
      TobTrav.PutValue('GA_DPA', dbTrav);
      TobTrav.PutValue('GA_PAHT', dbTrav);

      stDate := Trim(Copy(stEnreg, 426, 13));          // <424>  prix moyen pondere
      stSigne := Copy(stEnreg, 425, 1);
      iPos:=Pos('.', stDate);
      if iPos>0 then stDate[iPos] := ',';
      dbTrav := 0.0;
      if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
      if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
      TobTrav.PutValue('GA_PMAP', dbTrav);

      MessAvance:='Affectation minimum de marge #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stDate := Trim(Copy(stEnreg, 300, 13));          // <298>  Minimum de marge  999.99
      stSigne := Copy(stEnreg, 299, 1);
      iPos:=Pos('.', stDate);
      if iPos>0 then stDate[iPos] := ',';
      dbTrav := 0.0;
      if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
      if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
      TobTrav.PutValue('GA_MARGEMINI', dbTrav);

      MessAvance:='Affectation des poids brut, net, douaniers #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stDate := Trim(Copy(stEnreg, 468, 13));          // <466>  poids
      stSigne := Copy(stEnreg, 467, 1);
      iPos:=Pos('.', stDate);
      if iPos>0 then stDate[iPos] := ',';
      dbTrav := 0.0;
      if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
      if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
      TobTrav.PutValue('GA_POIDSNET', dbTrav);
      TobTrav.PutValue('GA_POIDSBRUT', dbTrav);
      TobTrav.PutValue('GA_POIDSDOUA', dbTrav);
      TobTrav.PutValue('GA_QUALIFPOIDS', 'K');

      MessAvance:='Affectation des volume et surface #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stDate := Trim(Copy(stEnreg, 496, 13));          // <494>  volume
      stSigne := Copy(stEnreg, 495, 1);
      iPos:=Pos('.', stDate);
      if iPos>0 then stDate[iPos] := ',';
      dbTrav := 0.0;
      if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
      if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
      TobTrav.PutValue('GA_VOLUME', dbTrav);
      TobTrav.PutValue('GA_QUALIFVOLUME', 'M3');

      stDate := Trim(Copy(stEnreg, 482, 13));          // <480>  surface
      stSigne := Copy(stEnreg, 481, 1);
      iPos:=Pos('.', stDate);
      if iPos>0 then stDate[iPos] := ',';
      dbTrav := 0.0;
      if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
      if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
      TobTrav.PutValue('GA_SURFACE', dbTrav);
      TobTrav.PutValue('GA_QUALIFSURFACE', 'M2');

      MessAvance:='Affectation de la quantité du prix de vente #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stDate := Trim(Copy(stEnreg, 314, 13));          // <312>  Quantité du prix de vente
      stSigne := Copy(stEnreg, 313, 1);
      iPos:=Pos('.', stDate);
      if iPos>0 then stDate[iPos] := ',';
      dbTrav := 0.0;
      if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
      if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
      TobTrav.PutValue('GA_PRIXPOURQTE', dbTrav);

      if TobTrav.GetValue('GA_TYPEARTICLE')='NOM' then
        begin
        MessAvance:='Affectation entête de la nomenclature #13 Enregistrement n° '+IntToStr(Nb_Lus);
        TOBNomen:=TOB.Create('NOMENENT', TobTrav, -1);
        TOBNomen.PutValue ('GNE_NOMENCLATURE', TOBTrav.GetValue('GA_CODEARTICLE'));
        TOBNomen.PutValue ('GNE_ARTICLE', TOBTrav.GetValue('GA_ARTICLE'));
        TOBNomen.PutValue ('GNE_LIBELLE', TOBTrav.GetValue('GA_LIBELLE'));
        TOBNomen.PutValue ('GNE_DATECREATION', TOBTrav.GetValue('GA_DATECREATION'));
        TOBNomen.PutValue ('GNE_DATEMODIF', V_PGI.DateEntree);
        TOBNomen.PutValue ('GNE_QTEDUDETAIL', 1);
        end;

      MessAvance:='Mise à jour dans la base #13 Enregistrement n° '+IntToStr(Nb_Lus);
      if ((Nb_Lus mod Ratio) = 0) then
          begin
          TobOut.SetAllModifie(True);
          if CB1.Checked then
              TobOut.InsertDB(nil)
              else
              TobOut.InsertOrUpdateDB;
          TobOut.ClearDetail;
          end;
      end;
  MessAvance:='Fin du traitement #13 Enregistrement n° '+IntToStr(Nb_Lus);
  CloseFile(FicSeq);
  TobOut.SetAllModifie(True);
  if CB1.Checked then
      TobOut.InsertDB(nil)
      else
      TobOut.InsertOrUpdateDB;
  TobOut.Free;
Except
  CloseFile(FicSeq);
  TobOut.Free;
  writeln (FicLog, 'Erreur : ' + MessAvance);
  MessAvance:=MessAvance+' #13 '+Copy(stEnreg,1,50);
  PGIInfo('Erreur : ' + MessAvance,'Récupération de dossier Négoce');
  OkOk:=False;
end;
end;

procedure TFRecupS3.Traite_Nomenclatures;
var
    stEnreg, stDate, stSigne, stTemp : string;
    FicSeq : textfile;
    TobOut, TobTrav : TOB;
    Nb_Lus, iDim, i_sizefic, Ratio, i_Trav, iPos : integer;
    dbTrav : double;
    Searchrec : TSearchRec;
    QQ : TQuery;
    ExistFic : boolean;
    MessAvance : string ;

    function LitEnreg : string;
    var  Ch : char;
         ind1 : integer;
    begin
    Result := '';
    for ind1 := 1 to LgNmcl do
        begin
        Read(FicSeq, Ch);
        Result := Result + Ch;
        end;
    end;

begin
if not FileExists('C:\PGI01\DAT\NmclS3.txt') then Exit;
SB1.SimpleText := 'Traitement de la table Nomenclature';
//  ouverture du fichier de données
{$I-}
AssignFile(FicSeq, 'C:\PGI01\DAT\NmclS3.txt');
Reset (FicSeq);
{$I+}
ExistFic:=(IOResult = 0);
if Not ExistFic then exit;
stEnreg := LitEnreg;
Nb_Lus := 0;
TobOut := TOB.Create('', nil, -1);
MessAvance:='Début de traitement Nomenclature ';
Try
  //  calcul nb enreg du fichier de données
  FindFirst('C:\PGI01\DAT\NmclS3.txt', faAnyFile, SearchRec);
  i_sizefic := Trunc(SearchRec.Size / Length(stEnreg));
  PB1.Position := 0;
  Ratio := Round(i_sizefic / 100);
  if Ratio = 0 then Ratio := 1;
  st_log := '****** Début du traitement NOMENCLATURE ******';
  writeln (FicLog, st_log);
  flush (FicLog);
  Reset (FicSeq);
  while not EOF(FicSeq) do
      begin
      stEnreg := LitEnreg;
  writeln (FicLog, Copy(stEnreg,1,50));
  flush (FicLog);
      Inc(Nb_Lus);
      if ((Nb_Lus mod Ratio) = 0) or (Ratio = 1) then PB1.StepIt;
      SB1.Panels.Items[1].Text := IntToStr(Nb_Lus) + ' sur ' + IntToStr(i_Sizefic);
      TobTrav := TOB.Create('NOMENLIG', TobOut, -1);
      TobTrav.InitValeurs;
  //  mise en place des valeurs importées
      MessAvance:='Affectation code nomenclature #13 Enregistrement n° '+IntToStr(Nb_Lus);
      if Copy(stEnreg, 65, 1) <>'A' then continue;   // <64> A==>Article, T==>Texte, F==>Fin
      stTemp:=Copy(stEnreg, 1, 15) ;             // <  0>  Code nomenclature (clé primaire)
      TobTrav.PutValue('GNL_NOMENCLATURE', Uppercase(stTemp));     // <  0>  Code nomenclature (clé primaire)
      RemplaceLesBadCaractere(stTEmp);
      iDim := 18 + (MaxDimension * 3);
      stDate := format('%-'+IntToStr(iDim)+'.'+IntToStr(iDim)+'sX',[Uppercase(stTemp)]);

      stTemp := 'Select GA_ARTICLE from ARTICLE where GA_ARTICLE="' + stDate + '" And GA_TYPEARTICLE="NOM"';
      if Not ExisteSQL(stTemp) then Continue;
      MessAvance:='Affectation code article et article #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stTemp:=Copy(stEnreg, 20, 15) ;                   // <19> Code article
      RemplaceLesBadCaractere(stTEmp);
      TobTrav.PutValue('GNL_CODEARTICLE'   ,Uppercase(stTemp) );
      iDim := 18 + (MaxDimension * 3);
      stDate := format('%-'+IntToStr(iDim)+'.'+IntToStr(iDim)+'sX',[Uppercase(stTemp)]);

      TobTrav.PutValue('GNL_ARTICLE', stDate);
      MessAvance:='Affectation du libellé article composant #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stTemp := 'Select GA_LIBELLE from ARTICLE where GA_ARTICLE="' + TobTrav.GetValue('GNL_ARTICLE') + '"';
      QQ:=OpenSQL(stTemp,True) ;
      if QQ.Eof then
        begin
        Ferme(QQ);
        Continue;
        end ;
      TobTrav.PutValue('GNL_LIBELLE', QQ.Fields[0].AsString);
      Ferme(QQ) ;
      MessAvance:='Affectation numéro de ligne du composant #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stDate:=Copy(stEnreg, 16, 4) ;          // <15> Numéro d'ordre    (clé primaire)
      if IsNumeric(stDate) then i_Trav := StrToInt(stDate) else i_trav:=0;
      TobTrav.PutValue('GNL_NUMLIGNE', i_Trav);
      MessAvance:='Affectation de la quantité du composant #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stDate:=Copy(stEnreg, 91, 13) ;         // <89> Quantité
      stSigne := Copy(stEnreg, 90, 1);
      iPos:=Pos('.', stDate);
      if iPos>0 then stDate[iPos] := ',';
      dbTrav := 0.0;
      if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
      if (dbTrav = 0.0) or (stSigne = '-') then continue ;
      TobTrav.PutValue('GNL_QTE',dbTrav) ;
      MessAvance:='Affectation des dates création et modification #13 Enregistrement n° '+IntToStr(Nb_Lus);
      TobTrav.PutValue('GNL_DATECREATION',V_PGI.DateEntree) ;
      TobTrav.PutValue('GNL_DATEMODIF',IDate1900) ;
      TobTrav.PutValue('GNL_CREATEUR',V_PGI.User) ;
      TobTrav.PutValue('GNL_UTILISATEUR',V_PGI.User) ;

      MessAvance:='Mise à jour dans la base #13 Enregistrement n° '+IntToStr(Nb_Lus);
      if ((Nb_Lus mod Ratio) = 0) then
          begin
          TobOut.SetAllModifie(True);
          if CB1.Checked then
              TobOut.InsertDB(nil)
              else
              TobOut.InsertOrUpdateDB;
          TobOut.ClearDetail;
          end;
      end;
  MessAvance:='Fin du traitement #13 Enregistrement n° '+IntToStr(Nb_Lus);
  CloseFile(FicSeq);
  TobOut.SetAllModifie(True);
  if CB1.Checked then
      TobOut.InsertDB(nil)
      else
      TobOut.InsertOrUpdateDB;
  TobOut.Free;
Except
  CloseFile(FicSeq);
  TobOut.Free;
  writeln (FicLog, 'Erreur : ' + MessAvance);
  MessAvance:=MessAvance+' #13 '+Copy(stEnreg,1,50);
  PGIInfo('Erreur : ' + MessAvance,'Récupération de dossier Négoce');
  OkOk:=False;
end;
end;


procedure TFRecupS3.Traite_Tarif(TypeTarf : string);
var
    stSelect, stEnreg, stDate, stSigne,FicS3,StTiers,CodeDebTiers : string;
    FicSeq : textfile;
    TobArticle, TobA, TobTiers, TobOut, TobTrav, TobTemp : TOB;
    Nb_Lus, i_sizefic, Ratio, iDim, LgTarif,iPos,MaxTarif : integer;
    dbTrav : double;
    Searchrec : TSearchRec;
    Q : TQuery;
    ExistFic : boolean;
    MessAvance : string ;


    function LitEnreg (LgEnreg : integer) : string;
    var  Ch : char;
         ind1 : integer;
    begin
    Result := '';
    for ind1 := 1 to LgEnreg do
        begin
        Read(FicSeq, Ch);
        Result := Result + Ch;
        end;
    end;

begin
if TypeTarf='CLI' then
  begin
  FicS3:='C:\PGI01\DAT\TarfS3.txt' ;
  StTiers:='Client';
  CodeDebTiers:=CodeClient;
  LgTarif:=LgTarf;
  end else
  begin
  FicS3:='C:\PGI01\DAT\TrfoS3.txt' ;
  StTiers:='Fournisseur';
  CodeDebTiers:=CodeFour;
  LgTarif:=LgTrfo;
  end;
if not FileExists(FicS3) then Exit;
{TobCC := TOB.Create('', nil, -1);
stEnreg := 'Select * from CHOIXCOD where CC_TYPE="TRC"';
Q := OpenSql(stEnreg, True);
TobCC.LoadDetailDB('CHOIXCOD', '', '', Q, False);
Ferme(Q); }
TobArticle := TOB.Create('', nil, -1);

TobTiers := TOB.Create('', nil, -1);
stEnreg := 'Select T_TIERS, T_DEVISE, T_TARIFTIERS from TIERS';
Q := OpenSql(stEnreg, True);
TobTiers.LoadDetailDB('TIERS', '', '', Q, False);
Ferme(Q);
SB1.SimpleText := 'Traitement de la table Tarif '+StTiers;
//  ouverture du fichier de données
{$I-}
AssignFile(FicSeq, FicS3);
Reset (FicSeq);
{$I+}
ExistFic:=(IOResult = 0);
if Not ExistFic then
  begin
  TobArticle.Free ; TobTiers.Free;
  exit;
  end;
stEnreg := LitEnreg (LgTarif);
Nb_Lus := 0; MaxTarif:=0;
TobOut := TOB.Create('', nil, -1);
MessAvance:='Début de traitement Tarifs '+StTiers;
Q:= OpenSql('Select Max(GF_TARIF) from TARIF ', True);
if not Q.Eof then MaxTarif:=Q.Fields[0].AsInteger;
Ferme(Q);
Try
  //  calcul nb enreg du fichier de données
  FindFirst(FicS3, faAnyFile, SearchRec);
  i_sizefic := Trunc(SearchRec.Size / Length(stEnreg));
  PB1.Position := 0;
  Ratio := Round(i_sizefic / 100);
  if Ratio = 0 then Ratio := 1;
  st_log := '****** Début du traitement TARIFS '+StTiers+' ******';
  writeln (FicLog, st_log);
  flush (FicLog);
  Reset (FicSeq);
  while not EOF(FicSeq) do
      begin
      stEnreg := LitEnreg(LgTarif);
      writeln (FicLog, Copy(stEnreg,1,50));
      flush (FicLog);
      Inc(Nb_Lus);
      if ((Nb_Lus mod Ratio) = 0) or (Ratio = 1) then PB1.StepIt;
      SB1.Panels.Items[1].Text := IntToStr(Nb_Lus) + ' sur ' + IntToStr(i_Sizefic);
      TobTrav := TOB.Create('TARIF', TobOut, -1);
      TobTrav.InitValeurs;
  //  mise en place des valeurs importées
      MessAvance:='Affectation numéro de tarif #13 Enregistrement n° '+IntToStr(Nb_Lus);
      TobTrav.PutValue('GF_TARIF', IntToStr(MaxTarif+Nb_Lus));

      TobTrav.PutValue('GF_TARIFTIERS', Trim(Copy(stEnreg, 2, 3)));      // <  1> code tarif                           //
      if TobTrav.GetValue('GF_TARIFTIERS') <> '' then
          begin
          TobTemp := TobCC.FindFirst(['CC_CODE'],[TobTrav.GetValue('GF_TARIFTIERS')], True);
          if TobTemp <> Nil then
              stEnreg := Copy(stEnreg, 1, 57) + Trim(TobTemp.GetValue('CC_LIBRE')) + Copy(stEnreg, 61, 255);
          end;
      MessAvance:='Affectation code tiers #13 Enregistrement n° '+IntToStr(Nb_Lus);
      if Trim(Copy(stEnreg, 5, 10)) <> '' then
          TobTrav.PutValue('GF_TIERS', CodeDebTiers + Trim(Copy(stEnreg, 5, 10)));     // <  4> code client                          //
      TobTrav.PutValue('GF_TARIFARTICLE', Trim(Copy(stEnreg, 15, 3)));     // < 14> code famille tarif                   //

      MessAvance:='Affectation code article #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stDate:=Copy(stEnreg, 18, 15);
      if Trim(stDate) <> '' then
          begin
          RemplaceLesBadCaractere(stDate);
          TobTrav.PutValue('GF_ARTICLE'   ,Uppercase(stDate) );        // < 17> code article
          iDim := 18 + (MaxDimension * 3);
          stDate := format('%-'+IntToStr(iDim)+'.'+IntToStr(iDim)+'sX',[TobTrav.GetValue('GF_ARTICLE')]);
          end;
      TobTrav.PutValue('GF_ARTICLE', stDate);

      MessAvance:='Affectation regime de prix et qualifiant #13 Enregistrement n° '+IntToStr(Nb_Lus);
      if Trim(TobTrav.GetValue('GF_ARTICLE')) <> '' then
          begin
          TobTrav.PutValue('GF_REGIMEPRIX', 'HT');
  //        TobTrav.PutValue('GF_REGIMEPRIX', 'HT')
          TobTemp := TobCC.FindFirst(['CC_CODE'],[TobTrav.GetValue('GF_TARIFTIERS')], True);
          if TobTemp <> Nil then
              TobTrav.PutValue('GF_REGIMEPRIX', TobTemp.GetValue('CC_ABREGE'))
              else
              begin
              TobTemp := TobTiers.FindFirst(['T_TIERS'],[TobTrav.GetValue('GF_TIERS')], True);
              if TobTemp <> Nil then
                  begin
                  TobTemp := TobCC.FindFirst(['CC_CODE'],[TobTemp.GetValue('T_TARIFTIERS')], True);
                  if TobTemp <> Nil then TobTrav.PutValue('GF_REGIMEPRIX', TobTemp.GetValue('CC_ABREGE'));
                  end;
              end;
          end
          else
          TobTrav.PutValue('GF_REGIMEPRIX', 'GLO');

      TobTrav.PutValue('GF_QUALIFPRIX', 'GRP');     // < 17> code article                         //
      TobTrav.PutValue('GF_CASCADEREMISE', 'MIE');
      MessAvance:='Affectation des bornes de quantité #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stDate := Trim(Copy(stEnreg, 34, 7));    // < 32> quantité plancher                    //
      stSigne := Copy(stEnreg, 33, 1);
      while Pos('.', stDate) <> 0 do stDate[Pos('.', stDate)] := ',';
      dbTrav := 0.0;
      if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
      if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
      if (dbTrav = 0.0) or (dbTrav = (-999999.0)) then
          begin
          TobTrav.PutValue('GF_BORNEINF', -999999);
          TobTrav.PutValue('GF_BORNESUP', 999999);
          TobTrav.PutValue('GF_QUANTITATIF', '-');
          end else
          begin
          TobTrav.PutValue('GF_BORNEINF', dbTrav);
          //if IsNumeric(Trim(Copy(stEnreg, 33, 8))) then
          //    TobTrav.PutValue('GF_BORNEINF', StrToFloat(Copy(stEnreg, 33, 8)));
          TobTrav.PutValue('GF_BORNESUP', 999999);
          TobTrav.PutValue('GF_QUANTITATIF', 'X');
          TobTrav.PutValue('GF_CASCADEREMISE', 'CAS');
          end;

      MessAvance:='Affectation du prix unitaire #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stDate := Trim(Copy(stEnreg, 100, 13));  // <120> prix unitaire de vente               //
      stSigne := Copy(stEnreg, 99, 1);
      iPos:=Pos('.', stDate);
      if iPos>0 then stDate[iPos] := ',';
      dbTrav := 0.0;
      if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
      if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
      TobTrav.PutValue('GF_PRIXUNITAIRE', dbTrav);

      stDate := Trim(Copy(stEnreg, 114, 13));          // <128> taux de coeff. ou de %               //
      stSigne := Copy(stEnreg, 113, 1);
      iPos:=Pos('.', stDate);
      if iPos>0 then stDate[iPos] := ',';
      dbTrav := 0.0;
      if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
      if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
  //    if Trim(Copy(stEnreg, 57, 1)) = 'C' then
  //        TobTrav.PutValue('GF_REMISE', (1 - dbTrav) * 100)
  //        else
          begin
          TobTrav.PutValue('GF_CALCULREMISE', stDate);
          TobTrav.PutValue('GF_REMISE', dbTrav);
          end;
      MessAvance:='Transformation coefficient en prix brut #13 Enregistrement n° '+IntToStr(Nb_Lus);
      if Trim(Copy(stEnreg, 57, 1)) = 'C' then
          begin
          if Trim(TobTrav.GetValue('GF_ARTICLE')) <> '' then
              begin
              TobA := TobArticle.FindFirst(['GA_ARTICLE'], [TobTrav.GetValue('GF_ARTICLE')], True);
              if TobA = nil then
                  begin
                  stSelect := 'Select GA_ARTICLE, GA_PVHT, GA_PVTTC from ARTICLE where GA_ARTICLE="' +
                              TobTrav.GetValue('GF_ARTICLE') + '"';
                  Q := OpenSql(stSelect, True);
                  TobArticle.LoadDetailDB('ARTICLE', '', '', Q, True);
                  Ferme(Q);
                  TobA := TobArticle.FindFirst(['GA_ARTICLE'], [TobTrav.GetValue('GF_ARTICLE')], True);
                  end;
              if TobA = nil then Continue;
              if TobTrav.GetValue('GF_REGIMEPRIX') = 'HT' then
                  TobTrav.PutValue('GF_PRIXUNITAIRE', TobA.GetValue('GA_PVHT') * TobTrav.GetValue('GF_REMISE'))
                  else
                  TobTrav.PutValue('GF_PRIXUNITAIRE', TobA.GetValue('GA_PVTTC') * TobTrav.GetValue('GF_REMISE'));
              TobTrav.PutValue('GF_REMISE', 0);
              TobTrav.PutValue('GF_CALCULREMISE', '');
              end
              else
              begin
              TobTrav.PutValue('GF_REMISE', (1 - TobTrav.GetValue('GF_REMISE')) * 100);
              TobTrav.PutValue('GF_CALCULREMISE', FloatToStr(TobTrav.GetValue('GF_REMISE')));
              end;
          end;

      MessAvance:='Affectation code devise du tarif #13 Enregistrement n° '+IntToStr(Nb_Lus);
      if Trim(Copy(stEnreg, 58, 3)) = '' then       // < 57> code devise                          //
          begin
          TobTemp := TobTiers.FindFirst(['T_TIERS'],[TobTrav.GetValue('GF_TIERS')], True);
          if TobTemp <> Nil then
              TobTrav.PutValue('GF_DEVISE', TobTemp.GetValue('T_DEVISE'))
              else
              TobTrav.PutValue('GF_DEVISE', 'EUR');
          end
          else
          TobTrav.PutValue('GF_DEVISE', Trim(Copy(stEnreg, 58, 3)));

      MessAvance:='Calcul de la priorité #13 Enregistrement n° '+IntToStr(Nb_Lus);
      if Trim(Copy(stEnreg, 2, 3)) <> '' then
          begin
          if Trim(Copy(stEnreg, 15, 3)) <> '' then
              TobTrav.PutValue('GF_PRIORITE', '17')
              else
              TobTrav.PutValue('GF_PRIORITE', '47');
          end
          else
          begin
          if Trim(Copy(stEnreg, 5, 10)) <> '' then
              begin
              if Trim(Copy(stEnreg, 15, 3)) <> '' then
                  TobTrav.PutValue('GF_PRIORITE', '52')
                  else
                  TobTrav.PutValue('GF_PRIORITE', '82');
              end
              else
              TobTrav.PutValue('GF_PRIORITE', '42');
          end;

      stDate := Copy(stEnreg, 2, 3) + ' - ' + Copy(stEnreg, 5, 10) + ' - ' +
                Copy(stEnreg, 15, 3) + ' - ' + Copy(stEnreg, 18, 15);
  {   if Trim(Copy(stEnreg, 5, 10)) <> '' then
          begin
          if Trim(Copy(stEnreg, 15, 3)) <> '' then
              TobTrav.PutValue('GF_LIBELLE', 'Client ' + Copy(stEnreg, 15, 3))
              else
              TobTrav.PutValue('GF_LIBELLE', 'Client ' + Copy(stEnreg, 15, 3) + ' Art ' + Copy(stEnreg, 18, 15));
          end
          else
          begin }
      TobTrav.PutValue('GF_LIBELLE', Trim(stDate));
      if TobTrav.GetValue('GF_QUANTITATIF') = 'X' then
          TobTrav.PutValue('GF_LIBELLE', 'Tarif Quantitatif');
  //      end;

      MessAvance:='Affectation des dates début, fin, effet #13 Enregistrement n° '+IntToStr(Nb_Lus);
      TobTrav.PutValue('GF_DATEEFFET', Date); // < 76> date effet nouveau PU                //
      TobTrav.PutValue('GF_DATEDEBUT', Date); // < 76> date effet nouveau PU                //
      TobTrav.PutValue('GF_DATEMODIF', Date); // < 76> date effet nouveau PU                //
      TobTrav.PutValue('GF_DATEFIN', iDate2099);
      TobTrav.PutValue('GF_NATUREAUXI', TypeTarf); // 'CLI' ou 'FOU'

      MessAvance:='Affectation ancien prix du tarif #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stDate := Trim(Copy(stEnreg, 86, 13));  // <112> ancien prix de vente                 //
      stSigne := Copy(stEnreg, 85, 1);
      iPos:=Pos('.', stDate);
      if iPos>0 then stDate[iPos] := ',';
      dbTrav := 0.0;
      if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
      if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
      TobTrav.PutValue('GF_PRIXANCIEN', dbTrav);

  //    if Trim(Copy(stEnreg, 57, 1)) = 'C' then TobTrav.Free;

  //
      MessAvance:='Mise à jour dans la base #13 Enregistrement n° '+IntToStr(Nb_Lus);
      if ((Nb_Lus mod Ratio) = 0) then
          begin
          TobOut.SetAllModifie(True);
          if CB1.Checked then
              TobOut.InsertDB(nil)
              else
              TobOut.InsertOrUpdateDB;
          TobOut.ClearDetail;
          end;
      end;
  MessAvance:='Fin du traitement #13 Enregistrement n° '+IntToStr(Nb_Lus);
  CloseFile(FicSeq);
  TobOut.SetAllModifie(True);
  if CB1.Checked then
      TobOut.InsertDB(nil)
      else
      TobOut.InsertOrUpdateDB;
  TobOut.Free;
Except
  CloseFile(FicSeq);
  TobOut.Free;
  writeln (FicLog, 'Erreur : ' + MessAvance);
  MessAvance:=MessAvance+' #13 '+Copy(stEnreg,1,50);
  PGIInfo('Erreur : ' + MessAvance,'Récupération de dossier Négoce');
  OkOk:=False;
end;
end;

procedure TFRecupS3.Traite_Tablettes;
var
    stEnreg, stDate, stSigne, stCode, stType, stSQL : string;
    FicSeq : textfile;
    TobMode, TobCX, TobPY, TobD, TobGCL, TobREG, TobTVA, TobTrav, TobTrav2 : TOB;
    Nb_Lus, i_sizefic, Ratio, ind1,iPos : integer;
    dbTrav : double;
    Searchrec : TSearchRec;
    Q : TQuery;
    ExistFic : boolean;
    MessAvance : string ;

    function LitEnreg : string;
    var  Ch : char;
         ind1 : integer;
    begin
    Result := '';
    for ind1 := 1 to LgPrm2 do
        begin
        Read(FicSeq, Ch);
        Result := Result + Ch;
        end;
    end;
    function LitEnreg2 : string;
    var  Ch : char;
         ind1 : integer;
    begin
    Result := '';
    for ind1 := 1 to 68 do
        begin
        Read(FicSeq, Ch);
        Result := Result + Ch;
        end;
    end;

begin
if not FileExists('C:\PGI01\DAT\Prm2S3.txt') then Exit;
SB1.SimpleText := 'Traitement des tablettes';
//  ouverture du fichier de données
{$I-}
AssignFile(FicSeq, 'C:\PGI01\DAT\Prm2S3.txt');
Reset (FicSeq);
{$I+}
ExistFic:=(IOResult = 0);
if Not ExistFic then exit;
stEnreg := LitEnreg;
Nb_Lus := 0;
TobCC  := TOB.Create('', nil, -1);
TobCX  := TOB.Create('', nil, -1);
TobPY  := TOB.Create('', nil, -1);
TobD   := TOB.Create('', nil, -1);
TobGCL := TOB.Create('', nil, -1);
TobREG := TOB.Create('', nil, -1);
TobTVA := TOB.Create('', nil, -1);
TobTrav := TOB.Create('DEVISE', TobD, -1);
TobTrav.PutValue('D_DEVISE', 'EUR');
TobTrav.PutValue('D_SYMBOLE', 'EUR');
TobTrav.PutValue('D_LIBELLE', 'EURO');
TobTrav.PutValue('D_DECIMALE', 2);
TobTrav.PutValue('D_QUOTITE', 1);
TobTrav.PutValue('D_PARITEEURO', 1);

//  calcul nb enreg du fichier de données
FindFirst('C:\PGI01\DAT\Prm2S3.txt', faAnyFile, SearchRec);
i_sizefic := Trunc(SearchRec.Size / Length(stEnreg));
PB1.Position := 0;
Ratio := Round(i_sizefic / 100);
if Ratio = 0 then Ratio := 1;
st_log := '****** Début du traitement PRM2S3 ******';
writeln (FicLog, st_log);
flush (FicLog);
Reset (FicSeq);
while not EOF(FicSeq) do
    begin
    stEnreg := LitEnreg;
writeln (FicLog, Copy(stEnreg,1,50));
flush (FicLog);
    Inc(Nb_Lus);
    if ((Nb_Lus mod Ratio) = 0) or (Ratio = 1) then PB1.StepIt;
    SB1.Panels.Items[1].Text := IntToStr(Nb_Lus) + ' sur ' + IntToStr(i_Sizefic);

    stCode := Copy(stEnreg, 1, 3);
    stType := Copy(stEnreg, 4, 3);

//  PAY --> PY (PAYS)
//  LNG --> LGU (CC)
//  REP --> GCL (COMMERCIAL)
//  DEV --> D (Table Devises)
//  SP1 --> FN1 (CC)
//  SC1 --> LT1 (ChoixExt)
//  TRF --> TRC (CC)
//  FAM --> TAR (CC)
//  TVA --> TVA (TXCPTTVA)
    if (stCode = 'SC1') or (stCode = 'SC2') or (stCode = 'SC3') then
        TobTrav := TOB.Create('CHOIXEXT', TobCX, -1)
    else if (stCode = 'SP1') or (stCode = 'SP2') or (stCode = 'SP3') or
            (stCode = 'LNG') or (stCode = 'FAM') or (stCode = 'TRF') then
        TobTrav := TOB.Create('CHOIXCOD', TobCC, -1)
    else if (stCode = 'PAY') then
        TobTrav := TOB.Create('PAYS', TobPY, -1)
    else if (stCode = 'DEV') then
        TobTrav := TOB.Create('DEVISE', TobD, -1)
    else if (stCode = 'REP') then
        TobTrav := TOB.Create('COMMERCIAL', TobGCL, -1)
    else if (stCode = 'REG') then
        TobTrav := TOB.Create('MODEPAIE', TobREG, -1)
    else if (stCode = 'TVA') then
        TobTrav := TOB.Create('TXCPTTVA', TobTVA, -1)
    else
        begin
//        stEnreg := LitEnreg;
        Continue;
        end;

    TobTrav.InitValeurs;
//  mise en place des valeurs importées
    if TobTrav.NomTable = 'CHOIXCOD' then
        begin
        if (stCode = 'LNG') then stCode := 'LGU'
        else if (stCode = 'SP1') then stCode := 'FN1'
        else if (stCode = 'SP2') then stCode := 'FN2'
        else if (stCode = 'SP3') then stCode := 'FN3'
        else if (stCode = 'FAM') then stCode := 'TAR'
        else if (stCode = 'TRF') then stCode := 'TRC';
        TobTrav.PutValue('CC_TYPE', stCode);
        TobTrav.PutValue('CC_CODE', stType);
        if (stCode = 'TRC') then
            begin
            TobTrav.PutValue('CC_LIBELLE', Copy(stEnreg, 7, 20));
            TobTrav.PutValue('CC_LIBRE', Copy(stEnreg, 36, 3));
            TobTrav.AddChampSup('TTC', True);
            if Copy(stEnreg, 27, 1) = 'T' then
                TobTrav.PutValue('TTC', 'TTC')
                else
                TobTrav.PutValue('TTC', 'HT');
            end
            else
            TobTrav.PutValue('CC_LIBELLE', Copy(stEnreg, 7, 25));
        end
    else if TobTrav.NomTable = 'CHOIXEXT' then
        begin
        if (stCode = 'SC1') then stCode := 'LT1'
        else if (stCode = 'SC2') then stCode := 'LT2'
        else if (stCode = 'SC3') then stCode := 'LT3';
        TobTrav.PutValue('YX_TYPE', stCode);
        TobTrav.PutValue('YX_CODE', stType);
        TobTrav.PutValue('YX_LIBELLE', Copy(stEnreg, 7, 25));
        end
    else if TobTrav.NomTable = 'PAYS' then
        begin
        TobTrav.PutValue('PY_PAYS', stType);
        TobTrav.PutValue('PY_LIBELLE', Copy(stEnreg, 7, 25));
        end
    else if TobTrav.NomTable = 'DEVISE' then
        begin
        TobTrav.PutValue('D_DEVISE', stType);
        TobTrav.PutValue('D_SYMBOLE', stType);
        TobTrav.PutValue('D_LIBELLE', Copy(stEnreg, 7, 25));
        TobTrav.PutValue('D_DECIMALE', 6);
        TobTrav.PutValue('D_QUOTITE', 1);
        stDate := Trim(Copy(stEnreg, 32, 9));          // <21> taux de change
        stSigne := Copy(stEnreg, 31, 1);
        iPos:=Pos('.', stDate);
        if iPos>0 then stDate[iPos] := ',';
        dbTrav := 0.0;
        if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
        if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
        TobTrav.PutValue('D_PARITEEURO', dbTrav);
        end
    else if TobTrav.NomTable = 'COMMERCIAL' then
        begin
        TobTrav.PutValue('GCL_TYPECOMMERCIAL', stCode);
        TobTrav.PutValue('GCL_COMMERCIAL', stType);
        TobTrav.PutValue('GCL_LIBELLE', Copy(stEnreg, 7, 25));
        stDate := Trim(Copy(stEnreg, 36, 5));          // <21> taux de change
        stSigne := Copy(stEnreg, 35, 1);
        iPos:=Pos('.', stDate);
        if iPos>0 then stDate[iPos] := ',';
        dbTrav := 0.0;
        if IsNumeric(stDate) then dbTrav := StrToFloat(stDate);
        if (dbTrav <> 0.0) and (stSigne = '-') then dbTrav := dbTrav * -1;
        TobTrav.PutValue('GCL_COMMISSION', dbTrav);
        end
    else if TobTrav.NomTable = 'MODEPAIE' then
        begin
        TobTrav.PutValue('MP_MODEPAIE', Copy(stEnreg, 4, 3));
        TobTrav.PutValue('MP_LIBELLE', Copy(stEnreg, 7, 25));
        end
    else if TobTrav.NomTable = 'TXCPTTVA' then
        begin
        TobTrav.Free;
        TobTrav := TOB.Create('CHOIXCOD', TobCC, -1);
        TobTrav.PutValue('CC_TYPE', 'TX1');
        TobTrav.PutValue('CC_CODE', stType);
        TobTrav.PutValue('CC_LIBELLE', Copy(stEnreg, 7, 20));
        TobTrav2 := TOB.Create('', nil, -1);
        stSQL := 'Select * from CHOIXCOD where CC_TYPE="RTV"';
        Q := OpenSQL(stSQL, True);
        TobTrav2.LoadDetailDB('CHOIXCOD', '', '', Q, False);
        Ferme(Q);
        for ind1 := 0 to TobTrav2.Detail.Count - 1 do
            begin
            TobTrav := TOB.Create('TXCPTTVA', TobREG, -1);
            TobTrav.PutValue('TV_TVAOUTPF', 'TX1');
            TobTrav.PutValue('TV_CODETAUX', stType);
            TobTrav.PutValue('TV_REGIME', TobTrav2.Detail[ind1].GetValue('CC_CODE'));
            end;
        end;
    end;
CloseFile(FicSeq);
TobCC.SetAllModifie(True);
TobCC.InsertOrUpdateDB;
TobPY.SetAllModifie(True);
TobPY.InsertOrUpdateDB;
TobPY.Free;
TobD.SetAllModifie(True);
TobD.InsertOrUpdateDB;
TobD.Free;
TobGCL.SetAllModifie(True);
TobGCL.InsertOrUpdateDB;
TobGCL.Free;
TobTVA.SetAllModifie(True);
TobTVA.InsertOrUpdateDB;
TobTVA.Free;
TobREG.SetAllModifie(True);
TobREG.InsertOrUpdateDB;

if not FileExists('C:\PGI01\DAT\ReglS3.txt') then Exit;
//  ouverture du fichier de données
AssignFile(FicSeq, 'C:\PGI01\DAT\ReglS3.txt');
Reset (FicSeq);
stEnreg := LitEnreg2;
Nb_Lus := 0;
MessAvance:='Début de traitement réglementd ';
Try
  TobMode  := TOB.Create('', nil, -1);
  //  calcul nb enreg du fichier de données
  FindFirst('C:\PGI01\DAT\ReglS3.txt', faAnyFile, SearchRec);
  i_sizefic := Trunc(SearchRec.Size / Length(stEnreg));
  PB1.Position := 0;
  Ratio := Round(i_sizefic / 100);
  if Ratio = 0 then Ratio := 1;
  st_log := '****** Début du traitement REGLS3 ******';
  writeln (FicLog, st_log);
  flush (FicLog);
  Reset (FicSeq);
  while not EOF(FicSeq) do
      begin
      stEnreg := LitEnreg2;
  writeln (FicLog, Copy(stEnreg,1,50));
  flush (FicLog);
      Inc(Nb_Lus);
      if ((Nb_Lus mod Ratio) = 0) or (Ratio = 1) then PB1.StepIt;
      SB1.Panels.Items[1].Text := IntToStr(Nb_Lus) + ' sur ' + IntToStr(i_Sizefic);

      MessAvance:='Affectation code réglement #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stCode := Copy(stEnreg, 4, 3);
      TobD := TobREG.FindFirst(['MP_MODEPAIE'], [stCode], True);
      TobTrav := TOB.Create('MODEREGL', TobMode, -1);
      TobTrav.InitValeurs;
      {$IFDEF TESTS1}
      TobTrav.PutValue('MR_MODEREGLE', Copy(stEnreg, 4, 3));
      {$ELSE}
      TobTrav.PutValue('MR_MODEREGLE', Copy(stEnreg, 1, 3));
      {$ENDIF}
      if TobD <> nil then TobTrav.PutValue('MR_LIBELLE', TobD.GetValue('MP_LIBELLE'));
      MessAvance:='Affectation nbre d''échéance et nbre de jours les séparant #13 Enregistrement n° '+IntToStr(Nb_Lus);
      TobTrav.PutValue('MR_TAUX1', 100);
      if stEnreg[7] = 'O' then
          TobTrav.PutValue('MR_APARTIRDE', 'FIN')
          else
          TobTrav.PutValue('MR_APARTIRDE', 'FAC');
      if Trim(Copy(stEnreg, 8, 1)) <> '' then
          TobTrav.PutValue('MR_NOMBREECHEANCE', StrToInt(Copy(stEnreg, 8, 1)));
      stCode := Copy(stEnreg, 9, 3);
      if Trim(stCode) <> '' then
          if StrToInt(stCode) = 7 then TobTrav.PutValue('MR_SEPAREPAR', 'SEM')
          else if StrToInt(stCode) = 15 then TobTrav.PutValue('MR_SEPAREPAR', 'QUI')
          else if StrToInt(stCode) = 30 then TobTrav.PutValue('MR_SEPAREPAR', '1M')
          else if StrToInt(stCode) = 60 then TobTrav.PutValue('MR_SEPAREPAR', '2M')
          else if StrToInt(stCode) = 90 then TobTrav.PutValue('MR_SEPAREPAR', '3M')
          else if StrToInt(stCode) = 120 then TobTrav.PutValue('MR_SEPAREPAR', '4M')
          else if StrToInt(stCode) = 150 then TobTrav.PutValue('MR_SEPAREPAR', '5M')
          else if StrToInt(stCode) = 180 then TobTrav.PutValue('MR_SEPAREPAR', '6M')
          else TobTrav.PutValue('MR_SEPAREPAR', '');
      TobTrav.PutValue('MR_MONTANTMIN', 999999);

      MessAvance:='Affectation quantième et nbre de jours #13 Enregistrement n° '+IntToStr(Nb_Lus);
      stDate := Trim(Copy(stEnreg, 41, 13));          // quantieme
      iPos:=Pos('.', stDate)-1;
      if iPos>0 then stDate := Trim(Copy(stDate, 0, iPos));
      if StrToInt(stDate) = 0 then
          TobTrav.PutValue('MR_ARRONDIJOUR', 'PAS')
          else
          TobTrav.PutValue('MR_ARRONDIJOUR', stDate + 'M');

      stDate := Trim(Copy(stEnreg, 32, 5));          // Nb jours
      if not IsNumeric(stDate) then stDate:='0';
      TobTrav.PutValue('MR_PLUSJOUR', StrToInt(stDate));
      if TobD <> nil then TobTrav.PutValue('MR_MP1', TobD.GetValue('MP_MODEPAIE'));

      end;
  MessAvance:='Mise à jour dans la base #13 Enregistrement n° '+IntToStr(Nb_Lus);
  CloseFile(FicSeq);
  TobMode.SetAllModifie(True);
  TobMode.InsertOrUpdateDB;
  TobMode.Free;
  TobREG.Free;
Except
  CloseFile(FicSeq);
  TobMode.Free; TobREG.Free;
  writeln (FicLog, 'Erreur : ' + MessAvance);
  MessAvance:=MessAvance+' #13 '+Copy(stEnreg,1,50);
  PGIInfo('Erreur : ' + MessAvance,'Récupération de dossier Négoce');
  OkOk:=False;
end;
end;

end.
