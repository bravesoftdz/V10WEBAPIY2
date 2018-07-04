unit factgrpBtp;

interface

Uses HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}  DB, Fe_Main,
      {$IFDEF V530} EdtDOC,{$ELSE} EdtRdoc, {$ENDIF}
{$ENDIF}
     UplannifChUtil,
     FactTOB,
     FactOuvrage,factspec,BTPUtil,
     FactRg,FactAcompte,FactComm,paramSoc,FactAdresse,
     SysUtils, Dialogs, SaisUtil, UtilPGI, AGLInit, FactUtil, UtilSais,
     EntGC, Classes, HMsgBox, ed_tools, UtilGrp,
     FactCpta, UtilArticle,uEntCommun,UtilTOBPiece,CalcOLEGenericBTP;

procedure AlimNewPieceBTP ( TOBGenere,TOBPiece : TOB);
procedure DefiniChampRuptBtp (var ChampsRupt : TStrings);
procedure DefiniTriEtRuptureBtp ( TOBSource : TOB ; var stSort : string);
procedure RegroupeLesAcomptesBTP (TOBPiece,TOBAcomptesGen : TOB);
function IsSameDocument(DocumentInit,DocumentInitPrec : string) : boolean;
function RecupPieceInit (TOBL : TOB) : string;
procedure AffecteAdresseLiv(TOBL: TOB;MemoTemp : Tstrings);
procedure InitLigneComm (TOBL,NewTobP,NewTOBL : TOB;iLigne : integer);
procedure ConstitueLienInterDoc (TOBgenere,TOBLiaison : TOB);
procedure getAdresseSoc (TOBL : TOB;MemoTemp : TStrings);
Procedure G_LigneCommentaireBTP ( TOBPiece,TOBL : TOB; ALaFin : boolean=false; Intervention:boolean=false) ; overload;
Procedure G_LigneCommentaireBTP ( TOBPiece : TOB; Cledoc : R_cledoc; ALaFin : boolean=false;Intervention:boolean=false) ; overload;
procedure G_AjouteOuvragesPiece (LesTOBOuvrages,TOBDESOUVRAGES ,TOBL,NewTOBP : TOB);
procedure G_RecupereDetailOuvrage(TOBL,TOBDESOUVRAGES,TOBOuvrage,TOBArticles : TOB ;var indiceOuv : Integer;RepriseAffaireRef : boolean; NewNature : string) ;
procedure RetrouvePieceFraisGrpBtp (var Cledoc2 : r_cledoc;TOBpiece : TOB; WithCreat : boolean=false);
Procedure G_LigneCommentaireNumBon ( TOBPiece,TOBL : TOB; ALaFin : boolean=false;Intervention:boolean=false ) ;

implementation
uses FactVariante,FactureBtp,UtilBTPgestChantier;

procedure DefiniChampRuptBtp (var ChampsRupt : TStrings);
begin
  with ChampsRupt do
  BEGIN
    Add('PIECE;AFF_GENERAUTO;=;')    ; Add('PIECE;GP_TIERS;=;');
    Add('PIECE;GP_AFFAIRE;=;')       ; Add('PIECE;GP_FACTUREHT;=;');
    Add('PIECE;GP_REGIMETAXE;=;')    ; Add('PIECE;GP_DEVISE;=;');
    Add('PIECE;GP_SAISIECONTRE;=;')  ; Add('PIECE;GP_ESCOMPTE;=;');
    Add('PIECE;RUPTMILLIEME;=;');
  END;
end;

procedure DefiniTriEtRuptureBtp ( TOBSource : TOB ; var stSort : string);
begin
  UG_AjouteLesChampsBtp (TobSource);
  StSort:='AFF_GENERAUTO;GP_TIERS;GP_AFFAIRE;GP_FACTUREHT;GP_REGIMETAXE;GP_DEVISE;GP_SAISIECONTRE;GP_ESCOMPTE;RUPTMILLIEME;GP_DATEPIECE;GP_NUMERO;' ;
end;

procedure G_AjouteOuvragesPiece (LesTOBOuvrages,TOBDESOUVRAGES ,TOBL,NewTOBP : TOB);
var LETOBDoc : TOB;
begin
  LeTOBDoc := TOBDESOUVRAGES.findFIrst(['NATUREPIECEG','SOUCHE','NUMERO'],
  																		 [
                                       TOBL.GetValue('GL_NATUREPIECEG'),
                                       TOBL.GetValue('GL_SOUCHE'),
                                       TOBL.GetValue('GL_NUMERO')
                                       ],true);
  if leTOBDoc <> nil then
  begin
    leTOBDoc.ChangeParent (LESTOBOUVRAGES,-1);
    LeTOBDoc.putValue('NATUREPIECEG',NewTOBP.GEtVAlue('GP_NATUREPIECEG'));
    LeTOBDoc.putValue('SOUCHE',NewTOBP.GEtVAlue('GP_SOUCHE'));
    LeTOBDoc.putValue('NUMERO',NewTOBP.GEtVAlue('GP_NUMERO'));
  end;
end;
procedure G_ChangeNaturePiece (TOBNewOuv: TOB; NewNature : string);
var indice : integer;
begin
	if TOBnewOUV.FieldExists ('BLO_NATUREPIECEG') then
  begin
  	TOBNEwOUV.PutValue('BLO_NATUREPIECEG',NewNature);
  end;
	for Indice := 0 to TOBNewOuv.detail.count -1 do
  begin
    G_ChangeNaturePiece (TOBNEwOUV.detail[Indice],NEwNature);
  end;
end;

procedure G_RecupereDetailOuvrage(TOBL,TOBDESOUVRAGES,TOBOuvrage,TOBArticles : TOB ;var indiceOuv : Integer;RepriseAffaireRef : boolean; NewNature : string) ;
var TOBOuvPiece : TOB;
		IndiceNomen : integer;
    TOBLIGOUV,TOBNewOUV : TOB;
    TypeArticle : string;
begin
  if not NaturepieceOKPourOuvrage (TOBL) then exit;
  TypeArticle := TOBL.GetValue('GL_TYPEARTICLE');
  if (TypeArticle<>'OUV') and (TypeArticle<>'ARP') then Exit ;
  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
  if IndiceNomen = 0 then exit;

  TOBOuvPiece := TOBDESOUVRAGES.findFirst(['NATUREPIECEG','SOUCHE','NUMERO'],
  																				[TOBL.GEtVAlue('GL_NATUREPIECEG'),
                                           TOBL.GetVAlue('GL_SOUCHE'),
                                           TOBL.GetValue('GL_NUMERO')],true);
  if TOBOuvPiece = nil then exit;
  TOBLIGOUv := TOBOuvPiece.detail[IndiceNomen-1];
  if TOBLigOuv <> nil then
  begin
  	TOBNewOUV := TOB.Create(' UNE LIGNE OUV',TOBOuvrage,-1);
    TOBNewOUV.Dupliquer (TOBLIGOUV,true,true);
    TOBL.PutValue('GL_INDICENOMEN',indiceOuv) ; Inc(indiceOuv) ;
    G_ChangeNaturePiece (TOBNewOuv, NewNature);
  end;
end;

procedure AlimNewPieceBTP ( TOBGenere,TOBPiece : TOB);
var valeur : string;
begin
	valeur := '';
  // On stocke le code affaire principale et le code affaire lié au devis
  TOBGenere.PutValue('GP_AFFAIRE', TOBPiece.GetValue('GP_AFFAIRE')) ;
  TOBGenere.PutValue('GP_AFFAIRE1',TOBPiece.GetValue('GP_AFFAIRE1')) ;
  TOBGenere.PutValue('GP_AFFAIRE2',TOBPiece.GetValue('GP_AFFAIRE2')) ;
  TOBGenere.PutValue('GP_AFFAIRE3',TOBPiece.GetValue('GP_AFFAIRE3')) ;
  TOBGenere.PutValue('GP_AVENANT', TOBPiece.GetValue('GP_AVENANT')) ;
  TOBGenere.PutValue('GP_AFFAIREDEVIS',TOBPiece.GetValue('GP_AFFAIREDEVIS')) ;
  TOBGenere.PutValue('GP_BLOCNOTE', TOBPiece.GetValue('GP_BLOCNOTE')) ;
	TOBGenere.AddChampSupValeur ('AFF_OKSIZERO',TOBPiece.GetValue('AFF_OKSIZERO'));
  if TOBPiece.fieldExists ('RUPTMILLIEME') then valeur := TOBPiece.GetValue('RUPTMILLIEME') else valeur := '';
	TOBGenere.AddChampSupValeur ('RUPTMILLIEME',Valeur);
  TOBGenere.PutValue('GP_FACTUREHT', TOBPiece.GetValue('GP_FACTUREHT')) ;
  TOBGenere.PutValue('GP_DEVISE', TOBPiece.GetValue('GP_DEVISE')) ;
  TOBGenere.PutValue('GP_TAUXDEV', TOBPiece.GetValue('GP_TAUXDEV')) ;
  TOBGenere.PutValue('GP_COTATION', TOBPiece.GetValue('GP_COTATION')) ;
  TOBGenere.PutValue('GP_DATETAUXDEV', TOBPiece.GetValue('GP_DATETAUXDEV')) ;
end;

procedure RegroupeLesAcomptesBTP (TOBPiece,TOBAcomptesGen : TOB);
var TOBACC,TOBLAC : TOB;
    Indice : integer;
begin
  if TOBPiece.detail.count > 0 then
  begin
    TOBAcc := TOB.Create ('LACOMPTE',TOBAcomptesGen,-1);
    AddLesChampSupAcptes (TOBAcc,TOBPiece);
    for Indice := 0 to TOBPiece.detail.count -1 do
    begin
      TOBLAC := TOB.Create ('ACOMPTES',TOBACC,-1);
      TOBLAC.Dupliquer (TOBPiece.detail[indice],true,true);
    end;
    TOBPiece.ClearDetail;
  end;
end;

function IsSameDocument(DocumentInit,DocumentInitPrec : string) : boolean;
var DocCur,DocPre : R_CleDoc;
begin
 if DocumentInitPrec = '' then begin result := false; exit; end;
 result := true;
 DecodeRefPiece (DocumentInit,DocCur);
 DecodeRefPiece (DocumentInitPrec,DocPre);
 if (DocPre.NaturePiece <> DocCur.NaturePiece) or
    (DocPre.NumeroPiece <> DocCur.NumeroPiece) or
    (DocPre.Indice <> DocCur.Indice) then result := false
end;

function RecupPieceInit (TOBL : TOB) : string;
begin
  Result := TOBL.GetValue('GL_PIECEORIGINE');
end;

procedure GetAdresseLivrClient (TOBPieceorig : TOB;MemoTemp :Tstrings);
var TOBAdresses : TOB;
		TOBContact : TOB;
    chaine : string;
begin
  TOBAdresses := TOB.Create('LESADRESSES', nil, -1);
  TOBCOntact := TOB.Create ('CONTACT',nil,-1);
  if GetParamSoc('SO_GCPIECEADRESSE') then
  begin
    TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Livraison}
    TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Facturation}
  end else
  begin
    TOB.Create('ADRESSES', TOBAdresses, -1); {Livraison}
    TOB.Create('ADRESSES', TOBAdresses, -1); {Facturation}
  end;
  TRY
    LoadLesAdresses(TOBPieceOrig, TOBAdresses);
    MemoTemp.add (TraduireMemoire('Livraison à l''adresse suivante'));
    if TOBAdresses.detail.count > 0 then
    begin
      if GetParamSoc('SO_GCPIECEADRESSE') then
      BEGIN
        if TobAdresses.detail[0].getvalue('GPA_LIBELLE') <> '' then
          MemoTemp.add (TobAdresses.detail[0].getvalue('GPA_LIBELLE'));
        if TobAdresses.detail[0].getvalue('GPA_LIBELLE2') <> '' then
          MemoTemp.add (TobAdresses.detail[0].getvalue('GPA_LIBELLE2'));
        if TobAdresses.detail[0].getvalue('GPA_ADRESSE1') <> '' then
          MemoTemp.add (TobAdresses.detail[0].getvalue('GPA_ADRESSE1'));
        if TobAdresses.detail[0].getvalue('GPA_ADRESSE2') <> '' then
          MemoTemp.add (TobAdresses.detail[0].getvalue('GPA_ADRESSE2'));
        if TobAdresses.detail[0].getvalue('GPA_ADRESSE3') <> '' then
          MemoTemp.add (TobAdresses.detail[0].getvalue('GPA_ADRESSE3'));
        if TobAdresses.detail[0].getvalue('GPA_CODEPOSTAL') <> '' then
          MemoTemp.add (TobAdresses.detail[0].getvalue('GPA_CODEPOSTAL')+ ' '+
                        TobAdresses.detail[0].getvalue('GPA_VILLE'));
        if TobAdresses.detail[0].getvalue('GPA_NUMEROCONTACT') <> 0 then
        begin
          if getContact (TOBContact,TOBPieceOrig.getValue('GP_TIERS'),TobAdresses.detail[0].getvalue('GPA_NUMEROCONTACT')) then
          begin
            chaine := 'Contact : '+rechDom('TTCIVILITE',TOBCOntact.getValue('C_CIVILITE'),false)  + ' '+
                      TOBContact.GetValue('C_PRENOM')+' '+ TOBContact.GetValue('C_NOM');
            if TOBContact.GetValue('C_TELEPHONE') <> '' then
            begin
              chaine := Chaine + ' '+Traduirememoire('Tel :')+' '+ TOBContact.GetValue('C_TELEPHONE');
            end;
            MemoTemp.add (chaine);
          end;
        end;
      END else
      begin
        if TobAdresses.detail[0].getvalue('ADR_libelle') <> '' then
          MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_libelle'));
        if TobAdresses.detail[0].getvalue('ADR_libelle2') <> '' then
          MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_libelle2'));
        if TobAdresses.detail[0].getvalue('ADR_ADRESSE1') <> '' then
          MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_ADRESSE1'));
        if TobAdresses.detail[0].getvalue('ADR_ADRESSE2') <> '' then
          MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_ADRESSE2'));
        if TobAdresses.detail[0].getvalue('ADR_ADRESSE3') <> '' then
          MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_ADRESSE3'));
        if TobAdresses.detail[0].getvalue('ADR_CODEPOSTAL') <> '' then
          MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_CODEPOSTAL')+' '+
                        TobAdresses.detail[0].getvalue('ADR_VILLE'));
        if TobAdresses.detail[0].getvalue('ADR_NUMEROCONTACT') <> 0 then
        begin
          if getContact (TOBContact,TOBPieceOrig.getValue('GP_TIERS'),TobAdresses.detail[0].getvalue('ADR_NUMEROCONTACT')) then
          begin
            chaine := 'Contact : '+rechDom('TTCIVILITE',TOBCOntact.getValue('C_CIVILITE'),false)  + ' '+
                      TOBContact.GetValue('C_PRENOM')+' '+ TOBContact.GetValue('C_NOM');
            if TOBContact.GetValue('C_TELEPHONE') <> '' then
            begin
              chaine := Chaine + ' '+Traduirememoire('Tel :')+' '+ TOBContact.GetValue('C_TELEPHONE');
            end;
            MemoTemp.add (chaine);
          end;
        end;
      end;
    end;
  FINALLY
    TOBAdresses.free;
    TOBCOntact.free;
  END;
end;

procedure getAdresseDepot ( Depot : string;MemoTemp : TStrings);
var TOBADD : TOB;
begin
  TOBADD := TOB.Create ('DEPOTS',nil,-1);
  TRY
    TOBADD.PutValue('GDE_DEPOT',Depot);
    TOBADD.LoadDB (true);
    MemoTemp.add (TraduireMemoire('Livraison au Dépot'));
    if TOBADD.getvalue('GDE_LIBELLE') <> '' then
      MemoTemp.add (TOBADD.getvalue('GDE_LIBELLE'));
    if TOBADD.getvalue('GDE_ADRESSE1') <> '' then
      MemoTemp.add (TOBADD.getvalue('GDE_ADRESSE1'));
    if TOBADD.getvalue('GDE_ADRESSE2') <> '' then
      MemoTemp.add (TOBADD.getvalue('GDE_ADRESSE2'));
    if TOBADD.getvalue('GDE_ADRESSE3') <> '' then
      MemoTemp.add (TOBADD.getvalue('GDE_ADRESSE3'));
    if TOBADD.getvalue('GDE_CODEPOSTAL') <> '' then
      MemoTemp.add (TOBADD.getvalue('GDE_CODEPOSTAL')+' '+
                    TOBADD.getvalue('GDE_VILLE'));
  FINALLY
    TOBADD.free;
  END;
end;

procedure getAdresseEtablissement ( Etablissement: string ;MemoTemp : Tstrings);
var TOBADD : TOB;
begin
  TOBADD := TOB.Create ('ETABLISS',nil,-1);
  TRY
    TOBADD.PutValue('ET_ETABLISSEMENT',Etablissement);
    TOBADD.LoadDB (true);
    MemoTemp.add (TraduireMemoire('Livraison à l''adresse suivante'));
    if TOBADD.getvalue('ET_LIBELLE') <> '' then
      MemoTemp.add (TOBADD.getvalue('ET_LIBELLE'));
    if TOBADD.getvalue('ET_ADRESSE1') <> '' then
      MemoTemp.add (TOBADD.getvalue('ET_ADRESSE1'));
    if TOBADD.getvalue('ET_ADRESSE2') <> '' then
      MemoTemp.add (TOBADD.getvalue('ET_ADRESSE2'));
    if TOBADD.getvalue('ET_ADRESSE3') <> '' then
      MemoTemp.add (TOBADD.getvalue('ET_ADRESSE3'));
    if TOBADD.getvalue('ET_CODEPOSTAL') <> '' then
      MemoTemp.add (TOBADD.getvalue('ET_CODEPOSTAL')+' '+
                    TOBADD.getvalue('ET_VILLE'));
  FINALLY
    TOBADD.free;
  END;
end;

procedure getAdresseSociete (MemoTemp : TStrings);
begin
  MemoTemp.add (TraduireMemoire('Livraison à l''adresse suivante'));
  if GetParamSoc('SO_LIBELLE') <> '' then
    MemoTemp.add (GetParamSoc('SO_LIBELLE'));
  if GetParamSoc('SO_ADRESSE1') <> '' then
    MemoTemp.add (GetParamSoc('SO_ADRESSE1'));
  if GetParamSoc('SO_ADRESSE2') <> '' then
    MemoTemp.add (GetParamSoc('SO_ADRESSE2'));
  if GetParamSoc('SO_ADRESSE3') <> '' then
    MemoTemp.add (GetParamSoc('SO_ADRESSE3'));
  if GetParamSoc('SO_CODEPOSTAL') <> '' then
    MemoTemp.add (GetParamSoc('SO_CODEPOSTAL')+' '+
                  GetParamSoc('SO_VILLE'));
end;

procedure getAdresseSoc (TOBL : TOB;MemoTemp : TStrings);
var Depot,Etablissement : String;
begin
  Depot := TOBL.GetValue('GL_DEPOT');
  Etablissement := TOBL.GetValue('GL_ETABLISSEMENT');
  if Depot <> '' then
  begin
    getAdresseDepot ( Depot,MemoTemp);
  end else if Etablissement <> '' then
  begin
    getAdresseEtablissement ( Etablissement,MemoTemp);
  end else
  begin
    getAdresseSociete (MemoTemp);
  end;
end;

procedure AffecteAdresseLiv(TOBL: TOB;MemoTemp : Tstrings);
var TOBPieceOrig : TOB;
   cledoc : R_CleDoc;
   Q : Tquery;
begin

  TOBPieceOrig := TOB.create ('PIECE',nil,-1);

  TRY
  	if TOBL.GetValue('GL_PIECEORIGINE') <> '' then
    begin
      DecoderefPiece (TOBL.GetValue('GL_PIECEORIGINE'),Cledoc);
      Q := OpenSQL('SELECT * FROM PIECE WHERE ' + WherePiece(CleDoc, ttdPiece, False), True,-1, '', True);
      if not Q.eof then TobPieceOrig.SelectDB('', Q);
      Ferme(Q);
    end;
    if TobL.getValue('GL_IDENTIFIANTWOL') = -1 then GetAdresseLivrClient (TOBPieceorig,MemoTemp)
                                               else getAdresseSoc (TOBL,memoTemp);
  FINALLY
    TobPieceOrig.free;
  END;

end;

procedure InitLigneComm (TOBL,NewTobP,NewTOBL : TOB;iLigne : integer);

  function GetReferenceOrigine (TOBL : TOB): string;
  var cledoc : r_cledoc;
      Q : Tquery;
      TheResult : string;
  begin
    Result := '';
    if TOBL.GetValue('GL_PIECEORIGINE') <> '' then
    begin
      DecoderefPiece (TOBL.GetValue('GL_PIECEORIGINE'),Cledoc);
      Q:=OpenSQL('SELECT GP_REFINTERNE FROM PIECE WHERE '+WherePiece(CleDoc,ttdPiece,False),True,-1, '', True) ;
      if not Q.eof then
      begin
        Theresult := Trim(Q.findField('GP_REFINTERNE').asString);
        if TheResult <> '' then Result := TraduireMemoire('N/Ref')+' : ' + TheResult;
      end;
      ferme (Q);
    end;
  end;

  function GetDocumentOrigine(TOBL : TOB ) : string;
  var cledoc : R_Cledoc;
  begin
    result := '';
    if TOBL.GetValue('GL_PIECEORIGINE') <> '' then
    begin
    	DecoderefPiece (TOBL.GetValue('GL_PIECEORIGINE'),Cledoc);
    end else Cledoc.NaturePiece := 'REA';

    if Cledoc.NaturePiece = 'REA' then
    begin
      result := TraduireMemoire('Commande stock');
    end else
    begin
      result := RechDom ('GCNATUREPIECEG',cledoc.NaturePiece ,false) + ' '+
                inttostr (cledoc.NumeroPiece);
    end;
  end;
var TheLibelle : string;
   theAffaire,TheReference : string;
begin
  PieceVersLigne(NewTOBP, NewTOBL);
  NewTOBL.PutValue('GL_NUMLIGNE', iLigne);
  NewTOBL.PutValue('RECALCULTARIF', '-');
  NewTOBL.PutValue('GL_PRIXPOURQTE', 1);
  NewTOBL.PutValue('GL_CODECOND', '');
  NewTOBL.PutValue('GL_INDICENOMEN', 0);
  NewTOBL.PutValue('GL_TYPELIGNE', 'RL');
  NewTOBL.PutValue('GL_TYPEDIM', 'NOR');
  NewTOBL.PutValue('GL_DATELIVRAISON', iDate1900);
  // Commissions
  TheAffaire :=TOBL.GetValue('GL_AFFAIRE');
  InitialiseComm(TOBL, False);
  NewTOBL.PutValue('GL_NUMORDRE', 0) ;
  NewTOBL.PutValue('GL_DEPOT', TOBL.GetValue('GL_DEPOT')) ;
  NewTOBL.putvalue('GL_AFFAIRE',TheAffaire);
  TheLibelle := GetDocumentOrigine(TOBL);
  TheReference := GetReferenceOrigine (TOBL);
  if TheReference <> '' then TheLibelle := TheLibelle + ' ' + TheReference;
  if TheAffaire <> '' then TheLibelle := TheLibelle + ' '+Traduirememoire('Chantier') +' '+ Copy(TheAffaire,2,14);
  NewTOBL.putvalue('GL_LIBELLE',TheLibelle);
end;

procedure DefiniLeLieninterDoc (TOBL,TOBLien : TOB);
var TOBLDD,TOBLI,TOBLD : TOB;
    Indice : integer;
begin
  TOBLD := TOB(TOBL.data); if TOBLD = nil then exit;
  for Indice := 0 to TOBLD.detail.count -1 do
  begin
    TOBLDD := TOBLD.detail[Indice];
    TOBLI := TOB.Create('LIENDEVCHA',TOBLien,-1);
    TOBLI.PutValue('BDA_REFD',EncodeLienDevCHA(TOBLDD));
    TOBLI.PutValue('BDA_NUMLD',TOBLDD.getValue('GL_NUMLIGNE'));
    if TOBLDD.FieldExists ('N1') then
    begin
      TOBLI.PutValue('BDA_N1D',TOBLDD.GetValue('N1'));
      TOBLI.PutValue('BDA_N2D',TOBLDD.GetValue('N2'));
      TOBLI.PutValue('BDA_N3D',TOBLDD.GetValue('N3'));
      TOBLI.PutValue('BDA_N4D',TOBLDD.GetValue('N4'));
      TOBLI.PutValue('BDA_N5D',TOBLDD.GetValue('N5'));
    end;
    TOBLI.PutValue('BDA_RANGD',Indice);
    TOBLI.PutValue('BDA_REFC',EncodeLienDevCHA(TOBL));
    TOBLI.PutValue('BDA_NUMLC',TOBL.GetValue('GL_NUMLIGNE'));
  end;
end;

procedure ConstitueLienInterDoc (TOBgenere,TOBLiaison : TOB);
var Indice : integer;
    TOBL{,TOBLP} : TOB;
begin
//  gcvoirtob(tobgenere);
  for Indice := 0 to TOBgenere.detail.count -1 do
  begin
     TOBL := TOBGenere.detail[Indice];
     if TOBL.data <> nil then DefiniLeLieninterDoc (TOBL,TOBLiaison);
  end;
end;

Procedure G_LigneCommentaireNumBon ( TOBPiece,TOBL : TOB; ALaFin : boolean=false;Intervention:boolean=false ) ;
Var NewL : TOB ;
    RefP : String ;
BEGIN
  //
  if TOBL.GetString('NUMBON')= '' then exit;
  //
  if AlAFin then NewL:=NewTOBLigne(TOBPiece,-1)
            else NewL:=NewTOBLigne(TOBPiece,1);
  NewL.PutValue('GL_NUMORDRE', 0) ;
  PieceVersLigne (TOBPiece,Newl);

  RefP:= 'Bon N° '+TOBL.GetString('NUMBON')+ ' du ' + DateToStr(TOBL.GetValue('GL_DATELIVRAISON'));
  RefP:=Copy(RefP,1,70) ;
  NewL.PutValue('GL_LIBELLE',RefP)    ; NewL.PutValue('GL_TYPELIGNE','COM') ;
  NewL.PutValue('GL_TYPEDIM','NOR')   ; NewL.PutValue('GL_CODEARTICLE','') ;
  NewL.PutValue('GL_ARTICLE','')      ; NewL.PutValue('GL_QTEFACT',0) ;
  NewL.PutValue('GL_QTESTOCK',0)      ; NewL.PutValue('GL_PUHTDEV',0) ;
  NewL.PutValue('GL_QTERESTE',0)      ; { NEWPIECE }
  //--- GUINIER ---
  NewL.PutValue('GL_MTRESTE',0)       ; {NEWPIECE}

  NewL.PutValue('GL_PUTTCDEV',0)      ; NewL.PutValue('GL_TYPEARTICLE','') ;
  NewL.PutValue('GL_PUHT',0)          ; NewL.PutValue('GL_PUHTNET',0) ;
  NewL.PutValue('GL_DPA',0)           ; NewL.PutValue('GL_DPR',0) ;
  NewL.PutValue('GL_PMAP',0)          ; NewL.PutValue('GL_PMRP',0) ;
  NewL.PutValue('GL_PUTTC',0)         ; NewL.PutValue('GL_PUTTCNET',0) ;
  NewL.PutValue('GL_PUHTBASE',0)      ; NewL.PutValue('GL_FAMILLETAXE1','') ;
  NewL.PutValue('GL_TYPENOMENC','')   ; NewL.PutValue('GL_QUALIFMVT','') ;
  NewL.PutValue('GL_REFARTSAISIE','') ; NewL.PutValue('GL_REFARTBARRE','') ;
  NewL.PutValue('GL_REFCATALOGUE','') ; NewL.PutValue('GL_TYPEREF','') ;
  NewL.PutValue('GL_REFARTTIERS','')  ;
  {Modif AC 4/07/03 Pas de GL_CODESDIM sur les lignes commentaire}
  NewL.PutValue('GL_CODESDIM','')  ;
  {Fin Modif AC}
  {Modif JLD 20/06/2002}
  NewL.PutValue('GL_ESCOMPTE',TOBPiece.GetValue('GP_ESCOMPTE')) ;
  NewL.PutValue('GL_REMISEPIED',TOBPiece.GetValue('GP_REMISEPIED')) ;
  {Fin modif}

  //JS 17/06/03
  NewL.PutValue('GL_INDICESERIE',0) ; NewL.PutValue('GL_INDICELOT',0) ;
  NewL.PutValue('GL_REMISELIGNE',0) ;
  // Modif BTP
  NewL.PutValue('GL_TYPEARTICLE','EPO');
  NewL.PutValue('GL_PUHTNETDEV',0)    ; NewL.PutValue('GL_PUTTCNETDEV',0) ;
  NewL.PutValue('GL_BLOCNOTE','')    ; NewL.PutValue('GL_QUALIFQTEVTE','') ;
  NewL.PutValue('GL_INDICENOMEN',0) ;
  NewL.PutValue('GL_NUMORDRE',0) ;
  NewL.PutValue('GL_VIVANTE','X') ;
  Newl.putvalue('GL_NIVEAUIMBRIC',0);
  NewL.PutValue('GL_MONTANTPA',0)  ;
  NewL.PutValue('GL_MONTANTPR',0)  ;
  NewL.PutValue('GL_MONTANTPAFG',0)  ;
  NewL.PutValue('GL_MONTANTPAFC',0)  ;
  NewL.PutValue('GL_MONTANTPAFR',0)  ;
  NewL.PutValue('GL_MONTANTFC',0)  ;
  NewL.PutValue('GL_MONTANTFG',0)  ;
  NewL.PutValue('GL_MONTANTFR',0)  ;

  // ---
  ZeroLigne(NewL) ;
END ;

Procedure G_LigneCommentaireBTP ( TOBPiece : TOB; Cledoc : R_cledoc; ALaFin : boolean=false; Intervention:boolean=false) ; overload;
Var NewL : TOB ;
    RefP : String ;
begin

	if AlAFin then
     NewL:=NewTOBLigne(TOBPiece,-1)
	else
     NewL:=NewTOBLigne(TOBPiece,1);

  NewL.PutValue('GL_NUMORDRE', 0);

  if Intervention then
     Begin
     RefP:= 'Intervention N° '+ BTPCodeAffaireAffiche(TOBPiece.getValue('GP_AFFAIRE')) + ' du ' + DateToStr(Cledoc.DatePiece);
     end
  else
     begin
     RefP:=RechDom('GCNATUREPIECEG',Cledoc.NaturePiece ,False)
          + ' N° '+IntToStr(Cledoc.NumeroPiece)
          + ' du '+DateToStr(Cledoc.DatePiece)
          + '  '  +TOBPiece.GetValue('GP_REFINTERNE'); //Attention ca vient d'un champs Sup issu de la requete
     end;

   RefP := Copy(RefP,1,70);
   NewL.PutValue('GL_LIBELLE',RefP);
   NewL.PutValue('GL_TYPELIGNE','COM');
   NewL.PutValue('GL_TYPEDIM','NOR');
   NewL.PutValue('GL_CODEARTICLE','');
   NewL.PutValue('GL_ARTICLE','');
   NewL.PutValue('GL_QTEFACT',0) ;
   NewL.PutValue('GL_QTESTOCK',0);
   NewL.PutValue('GL_PUHTDEV',0) ;
   NewL.PutValue('GL_DPA',0);
   NewL.PutValue('GL_DPR',0) ;
   NewL.PutValue('GL_PMAP',0);
   NewL.PutValue('GL_PMRP',0) ;
   NewL.PutValue('GL_QTERESTE',0); {NEWPIECE}
   //--- GUINIER ---
   NewL.PutValue('GL_MTRESTE',0); {NEWPIECE}
   NewL.PutValue('GL_PUTTCDEV',0);
   NewL.PutValue('GL_TYPEARTICLE','');
   NewL.PutValue('GL_PUHT',0);
   NewL.PutValue('GL_PUHTNET',0);
   NewL.PutValue('GL_PUTTC',0);
   NewL.PutValue('GL_PUTTCNET',0);
   NewL.PutValue('GL_PUHTBASE',0);
   NewL.PutValue('GL_FAMILLETAXE1','');
   NewL.PutValue('GL_TYPENOMENC','');
   NewL.PutValue('GL_QUALIFMVT','') ;
   NewL.PutValue('GL_REFARTSAISIE','');
   NewL.PutValue('GL_REFARTBARRE','');
   NewL.PutValue('GL_REFCATALOGUE','');
   NewL.PutValue('GL_TYPEREF','');
   NewL.PutValue('GL_REFARTTIERS','');
   NewL.PutValue('GL_MONTANTPA',0);
   NewL.PutValue('GL_MONTANTPR',0);
   NewL.PutValue('GL_MONTANTPAFG',0);
   NewL.PutValue('GL_MONTANTPAFC',0);
   NewL.PutValue('GL_MONTANTPAFR',0);
   NewL.PutValue('GL_MONTANTFC',0);
   NewL.PutValue('GL_MONTANTFG',0);
   NewL.PutValue('GL_MONTANTFR',0);

   {Modif AC 4/07/03 Pas de GL_CODESDIM sur les lignes commentaire}
   NewL.PutValue('GL_CODESDIM','');
   {Fin Modif AC}

   {Modif JLD 20/06/2002}
   NewL.PutValue('GL_ESCOMPTE', TOBPiece.GetValue('GP_ESCOMPTE')) ;
   NewL.PutValue('GL_REMISEPIED', TOBPiece.GetValue('GP_REMISEPIED')) ;
   {Fin modif}

   //JS 17/06/03
   NewL.PutValue('GL_INDICESERIE',0);
   NewL.PutValue('GL_INDICELOT',0);
   NewL.PutValue('GL_REMISELIGNE',0);

   // Modif BTP
   NewL.PutValue('GL_TYPEARTICLE','EPO');
   NewL.PutValue('GL_PUHTNETDEV',0);
   NewL.PutValue('GL_PUTTCNETDEV',0);
   NewL.PutValue('GL_BLOCNOTE','');
   NewL.PutValue('GL_QUALIFQTEVTE','');
   NewL.PutValue('GL_INDICENOMEN',0);
   NewL.PutValue('GL_NUMORDRE',0);
   NewL.PutValue('GL_VIVANTE','X');
   Newl.putvalue('GL_NIVEAUIMBRIC',0);
   PieceVersLigne (TOBPiece,Newl);

   // ---
   ZeroLigne(NewL) ;

end;

Procedure G_LigneCommentaireBTP ( TOBPiece,TOBL : TOB; ALaFin : boolean=false;Intervention:boolean=false ) ;
Var NewL : TOB ;
    RefP : String ;
BEGIN
if AlAFin then NewL:=NewTOBLigne(TOBPiece,-1)
					else NewL:=NewTOBLigne(TOBPiece,1);
(*
if TOBL<>Nil then
begin
	NewL.Dupliquer(TOBL,False,True) ;
  NewTOBLigneFille(Newl);
end;
*)
NewL.PutValue('GL_NUMORDRE', 0) ;
PieceVersLigne (TOBPiece,Newl);

if Intervention then
begin
  RefP:= 'Intervention N° '+BTPCodeAffaireAffiche(TOBL.getValue('GL_AFFAIRE'))
     +' du '+DateToStr(TOBL.GetValue('GL_DATELIVRAISON'));
end else
begin
  RefP:=RechDom('GCNATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG'),False)
     +' N° '+IntToStr(TOBL.GetValue('GL_NUMERO'))
     +' du '+DateToStr(TOBL.GetValue('GL_DATEPIECE'))
     +'  '+TOBL.GetValue('GP_REFINTERNE') ; // Attention ca vient d'un champs Sup issu de la requete
end;
RefP:=Copy(RefP,1,70) ;
NewL.PutValue('GL_LIBELLE',RefP)    ; NewL.PutValue('GL_TYPELIGNE','COM') ;
NewL.PutValue('GL_TYPEDIM','NOR')   ; NewL.PutValue('GL_CODEARTICLE','') ;
NewL.PutValue('GL_ARTICLE','')      ; NewL.PutValue('GL_QTEFACT',0) ;
NewL.PutValue('GL_QTESTOCK',0)      ; NewL.PutValue('GL_PUHTDEV',0) ;
NewL.PutValue('GL_QTERESTE',0)      ; { NEWPIECE }
//--- GUINIER ---
NewL.PutValue('GL_MTRESTE',0)       ; {NEWPIECE}

NewL.PutValue('GL_PUTTCDEV',0)      ; NewL.PutValue('GL_TYPEARTICLE','') ;
NewL.PutValue('GL_PUHT',0)          ; NewL.PutValue('GL_PUHTNET',0) ;
NewL.PutValue('GL_DPA',0)           ; NewL.PutValue('GL_DPR',0) ;
NewL.PutValue('GL_PMAP',0)          ; NewL.PutValue('GL_PMRP',0) ;
NewL.PutValue('GL_PUTTC',0)         ; NewL.PutValue('GL_PUTTCNET',0) ;
NewL.PutValue('GL_PUHTBASE',0)      ; NewL.PutValue('GL_FAMILLETAXE1','') ;
NewL.PutValue('GL_TYPENOMENC','')   ; NewL.PutValue('GL_QUALIFMVT','') ;
NewL.PutValue('GL_REFARTSAISIE','') ; NewL.PutValue('GL_REFARTBARRE','') ;
NewL.PutValue('GL_REFCATALOGUE','') ; NewL.PutValue('GL_TYPEREF','') ;
NewL.PutValue('GL_REFARTTIERS','')  ;
{Modif AC 4/07/03 Pas de GL_CODESDIM sur les lignes commentaire}
NewL.PutValue('GL_CODESDIM','')  ;
{Fin Modif AC}
{Modif JLD 20/06/2002}
NewL.PutValue('GL_ESCOMPTE',TOBPiece.GetValue('GP_ESCOMPTE')) ;
NewL.PutValue('GL_REMISEPIED',TOBPiece.GetValue('GP_REMISEPIED')) ;
{Fin modif}

//JS 17/06/03
NewL.PutValue('GL_INDICESERIE',0) ; NewL.PutValue('GL_INDICELOT',0) ;
NewL.PutValue('GL_REMISELIGNE',0) ;
// Modif BTP
NewL.PutValue('GL_TYPEARTICLE','EPO');
NewL.PutValue('GL_PUHTNETDEV',0)    ; NewL.PutValue('GL_PUTTCNETDEV',0) ;
NewL.PutValue('GL_BLOCNOTE','')    ; NewL.PutValue('GL_QUALIFQTEVTE','') ;
NewL.PutValue('GL_INDICENOMEN',0) ;
NewL.PutValue('GL_NUMORDRE',0) ;
NewL.PutValue('GL_VIVANTE','X') ;
Newl.putvalue('GL_NIVEAUIMBRIC',0);
NewL.PutValue('GL_MONTANTPA',0)  ;
NewL.PutValue('GL_MONTANTPR',0)  ;
NewL.PutValue('GL_MONTANTPAFG',0)  ;
NewL.PutValue('GL_MONTANTPAFC',0)  ;
NewL.PutValue('GL_MONTANTPAFR',0)  ;
NewL.PutValue('GL_MONTANTFC',0)  ;
NewL.PutValue('GL_MONTANTFG',0)  ;
NewL.PutValue('GL_MONTANTFR',0)  ;

// ---
ZeroLigne(NewL) ;
END ;


procedure GetPiecePrec (TOBPiece : TOB;var cledoc : r_cledoc);
var indice : integer;
    TOBL,TOBPiecePrec : TOB;
    QQ,QNext: Tquery;
    cledoc2,clesuiv : r_cledoc;
    PiecePrec : string;
    PieceFrais : string;
begin
  FillChar(CleDoc2, Sizeof(CleDoc2), #0);

  for indice := 0 to TOBPiece.detail.count -1 do
  begin
    TOBL := TOBPiece.detail[Indice];
    if TOBL.getValue('GL_PIECEORIGINE') <> '' then
    begin
      DecodeRefPiece (TOBL.getValue('GL_PIECEORIGINE'),cledoc2);
      break;
    end;
  end;

  if cledoc2.numeroPiece <> 0 then
  begin
  	cledoc := cledoc2;

  	cledoc.NaturePiece := 'FRC';
    if not ExistePiece(CleDoc) then
    begin
    	// retrour sur piece precedente
      QQ := OpenSql ('SELECT GP_PIECEFRAIS FROM PIECE WHERE  '+WherePiece(cledoc2,ttdPiece,false),true,-1, '', True);
      if not QQ.eof then
      begin
      	pieceFrais := QQ.findField('GP_PIECEFRAIS').asString;
        if PieceFrais = '' then
        begin
        	TOBPiecePrec := TOB.Create ('LES LIGNES',nil,-1);
          clesuiv := cledoc2 ;
      		QNext := OpenSql ('SELECT GL_PIECEORIGINE FROM LIGNE WHERE  '+WherePiece(clesuiv,ttdLigne,false)+ ' AND GL_PIECEORIGINE <> ""',true,-1, '', True);
          TOBPiecePrec.LoadDetailDB ('LIGNE','','',QNext,false);
          ferme (Qnext);
          GetPiecePrec (TOBPiecePrec,cledoc);
          TOBpieceprec.free;
          if cledoc.NumeroPiece = 0 then
          begin
          	FillChar(CleDoc, Sizeof(CleDoc), #0);
          end;
        end else
        begin
          DecodeRefPiece (PieceFrais,cledoc);
        end;
      end else
      begin
        FillChar(CleDoc, Sizeof(CleDoc), #0);
      end;
      ferme (QQ);
    end;
  end;
end;

procedure RetrouvePieceFraisGrpBtp (var Cledoc2 : r_cledoc;TOBpiece : TOB; WithCreat : boolean=false);
var TOBP : TOB;
		EnHt ,SaisieContre: boolean;
begin
	if TOBPiece.getValue ('GP_PIECEFRAIS')='' then
  begin
    FillChar(CleDoc2, Sizeof(CleDoc2), #0);
    GetPiecePrec (TOBPiece,cledoc2);
    if not ExistePiece(CleDoc2) then
    begin
    	// va bien falloir la créer cette piece
      // on reinit les numero de piece
      cledoc2.Souche := '';
      CleDoc2.NumeroPiece := 0;
      CleDoc2.Indice := 0;
      if WithCreat then
      begin
        EnHt := TOBPiece.GetValue('GP_FACTUREHT') = 'X';
        SaisieContre := TOBPiece.GetValue('GP_SAISIECONTRE') = 'X';
        CreerPieceVide(CleDoc2, TOBPiece.GetValue('GP_TIERS'), TOBPiece.GetValue('GP_AFFAIRE'), TOBPiece.GetValue('GP_ETABLISSEMENT'),
          TOBPiece.GetValue('GP_DOMAINE'), EnHT, SaisieContre);
      	TOBP := TOB.Create ('PIECE',nil,-1);
        PositionnePiece(TOBP,Cledoc2);
        TOBPiece.putValue('GP_PIECEFRAIS',EncodeRefPiece (TOBP));
        TOBPiece.putValue('_NEWFRAIS_','X');
        TOBP.free;
      end else
      begin
      	FillChar(CleDoc2, Sizeof(CleDoc2), #0);
      end;
    end else
    begin
    	// gestion antériorité
    	exit; // elle est rerouvéeee ouhééé champomy pour tout le monde...
    end;
  end else
  begin
  	DecodeRefPiece (TOBpiece.getValue('GP_PIECEFRAIS'),cledoc2);
    if (not ExistePiece(CleDoc2)) then
    begin
    	if WithCreat then
      begin
      	CleDoc2.NumeroPiece := 0;
        EnHt := TOBPiece.GetValue('GP_FACTUREHT') = 'X';
        SaisieContre := TOBPiece.GetValue('GP_SAISIECONTRE') = 'X';
        CreerPieceVide(CleDoc2, TOBPiece.GetValue('GP_TIERS'), TOBPiece.GetValue('GP_AFFAIRE'), TOBPiece.GetValue('GP_ETABLISSEMENT'),
          TOBPiece.GetValue('GP_DOMAINE'), EnHT, SaisieContre);
      	TOBP := TOB.Create ('PIECE',nil,-1);
        PositionnePiece(TOBP,Cledoc2);
        TOBPiece.putValue('GP_PIECEFRAIS',EncodeRefPiece (TOBP));
        TOBPiece.putValue('_NEWFRAIS_','X');
        TOBP.free;
      end else
      begin
    		FillChar(CleDoc2, Sizeof(CleDoc2), #0);
      end;
    end;
  end;
end;

end.
