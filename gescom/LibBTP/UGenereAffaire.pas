unit UGenereAffaire;

interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     FE_Main,
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     Maineagl,
     eMul,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
{$IFDEF BTP}
     CalcOLEGenericBTP,
{$ENDIF}
     UTob,
     ParamSoc,
     Splash,
     HMsgBox;

procedure GenereAffaireFromAnal;

implementation

procedure EcritAffaires (TOBAffaires : TOB);
begin
  TOBAFFaires.InsertDB (nil,true);
end;

procedure CreerAffaire (TOBAFFAIRES,TOBSECTION,TOBTIER : TOB);
var TOBAFF : TOB;
    Part,Part0,PArt1,PArt2,PART3,PART4 : string;
    DateTmp : TDateTime;
begin

  TOBAFF := TOB.Create ('AFFAIRE',TOBAFFAIRES,-1);

  Part := 'A'+TOBSECTION.GetValue('S_SECTION');
  Part := Part + Copy ('                 ',1,17-Length(Part));
  Part := Copy (Part,1,15)+'00';
  TOBAFF.putValue('AFF_AFFAIRE',part);
  TOBAFF.putValue('AFF_AFFAIREREF',Part);
  TOBAFF.putValue('AFF_AFFAIREINIT',Part);
  TOBAFF.putValue('AFF_LIBELLE',TOBSECTION.GetValue('S_LIBELLE'));
  Part0 := '';
  Part1 := '';
  Part2 := '';
  Part3 := '';
  Part4 := '';
  // Découpage du code Affaire
  {$IFDEF BTP}
  BTPCodeAffaireDecoupe(Part,Part0,Part1,Part2,Part3, Part4, TaModif,false);  
  {$ELSE}
  CodeAffaireDecoupe(Part,Part0,Part1,Part2,Part3, Part4, TaModif,false);
  {$ENDIF}
  TOBAFF.putValue('AFF_AFFAIRE0',Part0);
  TOBAFF.putValue('AFF_AFFAIRE1',Part1);
  TOBAFF.putValue('AFF_AFFAIRE2',Part2);
  TOBAFF.putValue('AFF_AFFAIRE3',Part3);
  TOBAFF.putValue('AFF_AVENANT',Part4);
  TOBAFF.putValue('AFF_TIERS',TOBTier.getValue('T_TIERS'));
  TOBAFF.putValue('AFF_STATUTAFFAIRE','AFF');
  TOBAFF.putValue('AFF_ETATAFFAIRE','ACP');

  TOBAFF.putValue('AFF_CREATEUR',V_PGI.User);
  TOBAFF.putValue('AFF_UTILISATEUR',V_PGI.User);
  TOBAFF.putValue('AFF_CREERPAR',V_PGI.User);

  TOBAFF.putValue('AFF_DATECREATION',V_PGI.DateEntree);
  TOBAFF.putValue('AFF_DATEMODIF',V_PGI.DateEntree);
  TOBAFF.putValue('AFF_DATEMODIF',V_PGI.DateEntree);
  TOBAFF.putValue('AFF_DATEREPONSE',V_PGI.dateEntree);
  TOBAFF.putValue('AFF_DATESIGNE',V_PGI.dateEntree);
  TOBAFF.putValue('AFF_DATEDEBUT',V_PGI.dateEntree);
  DateTmp := EncodeDate (2099,12,31);
  TOBAFF.putValue('AFF_DATEFIN',DateTmp);
  TOBAFF.putValue('AFF_DATELIMITE',DateTmp);
  TOBAFF.putValue('AFF_DATECLOTTECH',DateTmp);
  TOBAFF.putValue('AFF_DATEGARANTIE',DateTmp);
  TOBAFF.putValue('AFF_DATERESIL',DateTmp);
  TOBAFF.putValue('AFF_DATEFINEXER',DateTmp);
  TOBAFF.putValue('AFF_DATEFINGENER',DateTmp);
  DateTmp := EncodeDate (1900,01,01);
  TOBAFF.putValue('AFF_DATEDEBEXER',DateTmp);
  TOBAFF.putValue('AFF_DATFACECLATMOD',DateTmp);
  TOBAFF.putValue('AFF_DATESITUATION',DateTmp);
  TOBAFF.putValue('AFF_DATEDEBGENER',DateTmp);
  //
  TOBAFF.putValue('AFF_DEVISE',TOBTier.GetValue('T_DEVISE'));
  TOBAFF.putValue('AFF_AFFAIREHT',TOBTier.GetValue('T_FACTUREHT'));
  //
  TOBAFF.putValue('AFF_GENERAUTO',GetParamSoc('SO_AFFGENERAUTO'));
  TOBAFF.putValue('AFF_PERIODICITE','M');
  TOBAFF.putValue('AFF_COEFFREVALO',1);
  TOBAFF.putValue('AFF_RECONDUCTION','TAC');
  TOBAFF.putValue('AFF_REGROUPEFACT','AUC');
  TOBAFF.putValue('AFF_REPRISEACTIV','NON');
  TOBAFF.putValue('AFF_CONFIDENTIEL','0');
  TOBAFF.putValue('AFF_AFFCOMPLETE','X');
end;

procedure LanceTraitementCreatAff;
var TOBANAL,TOBS,TOBAFFAIRES,TOBTiers : TOB;
    QQ,Q1 : Tquery;
    Indice  : integer;
    Tiers : string;
begin
  if GetParamSoc('SO_GCTIERSDEFAUT') = '' then
  begin
    PGIBox (TraduireMemoire('Vous devez paramétrer le Tiers des transferts'),
            TraduireMemoire('Generation affaire'));
    Exit;
  end;

  Tiers := GetParamSoc('SO_GCTIERSDEFAUT');
  TOBTiers := TOB.Create ('TIERS',nil,-1);
  QQ := OpenSql ('SELECT * FROM TIERS WHERE T_TIERS="'+Tiers+'"',True,-1,'',true);
  if QQ.eof then
  BEGIN
    Ferme (QQ);
    PGIBox (TraduireMemoire('Le tiers des transferts est inexistant'),
            TraduireMemoire('Generation affaire'));
    Exit;
  END;

  TOBTiers.SelectDB ('',QQ);
  QQ := OPenSql ('SELECT * FROM DECOUPEANA WHERE GDA_AXE="A1"',true,-1,'',true);
  if QQ.eof then
  begin
    Ferme (QQ);
    PGIBox (TraduireMemoire('Vous devez paramétrer votre analytique'),
            TraduireMemoire('Generation affaire'));
    Exit;
  end;
  Ferme(QQ);

  TOBANAL := TOB.Create ('LES VENTILS ANAL',nil,-1);
  TOBAFFAIRES := TOB.Create ('LES AFFAIRES',nil,-1);
  QQ := OpenSql ('SELECT * FROM SECTION WHERE S_AXE="A1" AND S_FERME<>"X"',True,-1,'',true);
  TRY
    if not QQ.eof then
    begin
      TOBANAL.LoadDetailDB ('SECTION','','',QQ,true);
      for indice := 0 to TOBANAL.detail.count -1 do
      begin
        TOBS := TOBANAL.Detail[Indice];
        Q1 := OpenSql ('SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE="A'+TOBS.GetValue('S_SECTION')+'"',true,-1,'',true);
        if Q1.eof then
        begin
          CreerAffaire (TOBAFFAIRES,TOBS,TOBTiers);
        end;
        ferme (Q1);
      end;
    end;
    EcritAffaires (TOBAFFAIRES);
  FINALLY
    Ferme (QQ);
    TOBANAL.free;
    TOBAFFAIRES.Free;
  END;
end;


procedure GenereAffaireFromAnal;
var  splashScreen: TFsplashScreen;
begin
  if PGIAsk (TraduireMemoire('Désirez-vous générer les affaires à partir de l''analytique'),
             TraduireMemoire('Generation affaire')) = Mryes then
  begin
    splashScreen := TFsplashScreen.Create (application);
    splashScreen.Caption := TraduireMemoire('Generation affaire');
    splashScreen.Label1.Caption := TraduireMemoire('Génération d''affaires');
    splashScreen.Show;
    splashScreen.Update;
    LanceTraitementCreatAff;
    splashScreen.free;
  end;
end;

end.
