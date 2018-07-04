{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 28/06/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MBOEPURATIONALF ()
Mots clefs ... : TOF;MBOEPURATIONALF
*****************************************************************}
Unit MBOEPURATIONALF_TOF ;

Interface

Uses UTOF, HQry, ComCtrls,
{$IFDEF EAGLCLIENT}
     MaineAGL, eMul,
{$ELSE}
     Fe_Main, Mul,
{$ENDIF}
     HEnt1, HMsgBox, Controls, HStatus, EntGC, HCtrls, FactUtil, Forms,
     Classes, M3FP,
     FactTOB ;

Type
  TOF_MBOEPURATIONALF = Class (TOF)
  Private
    StWhere : string;
    procedure BatchEPURATION_ALF;
    procedure SetEPURATION;
    procedure SetAllEPURATION;
  Public
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

Procedure LanceEpurationAnnonceDeLivraison;

Implementation

Procedure LanceEpurationAnnonceDeLivraison;
begin
  AGLLanceFiche('MBO','EPURATION_ALF','GP_NATUREPIECEG=ALF','','');
end;

procedure TOF_MBOEPURATIONALF.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_MBOEPURATIONALF.OnLoad ;
begin
  Inherited ;
  stWhere := RecupWhereCritere(TPageControl(TFMul(Ecran).Pages));
end ;

procedure TOF_MBOEPURATIONALF.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_MBOEPURATIONALF.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_MBOEPURATIONALF.BatchEPURATION_ALF;
var i : integer;
begin
  with TFMul(Ecran) do
   begin
   if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
     begin PGIInfo('Veuillez sélectionner les pièces à épurer', Caption); exit; end;
   SourisSablier;
   try
     if FListe.AllSelected then
       begin
       if PGIAsk('Voulez-vous épurer toutes les pièces ?', Caption) <> mrYes then exit;
       if Transactions(SetAllEPURATION,2) <> oeOK then PGIError('Impossible d''épurer toutes les pièces', Caption);
       FListe.AllSelected := false;
       end
     else
       begin
       if PGIAsk('Voulez-vous épurer les pièces sélectionnées ?', Caption) <> mrYes then exit;
       InitMove(FListe.NbSelected,'');
       For i := 0 to FListe.NbSelected-1 do
         begin
         FListe.GotoLeBookMark(i);
         {$IFDEF EAGLCLIENT}
           Q.TQ.Seek(FListe.Row-1) ;
         {$ENDIF}
         if Transactions(SetEPURATION,2) <> oeOK then PGIBox('Impossible d''épurer la pièce n° '+Q.FindField('GP_NUMERO').AsString, Caption);
         MoveCur(False);
         end;
       FListe.ClearSelected;
       FiniMove;
       end;
     ChercheClick;
   Finally
    bSelectAll.Down := False;
    SourisNormale;
    end;
   end;
end ;

procedure TOF_MBOEPURATIONALF.SetEPURATION;
var DateSupp : String;
    CleDoc : R_CleDoc;
begin
  DateSupp := UsTime(NowH);
  with TFMul(Ecran) do
    begin
    CleDoc.NaturePiece := Q.FindField('GP_NATUREPIECEG').AsString;
    CleDoc.DatePiece := Q.FindField('GP_DATEPIECE').AsDateTime;
    CleDoc.Souche := Q.FindField('GP_SOUCHE').AsString;
    CleDoc.NumeroPiece := Q.FindField('GP_NUMERO').AsInteger;
    CleDoc.Indice := Q.FindField('GP_INDICEG').AsInteger;
    end;
  //Tag la pièce supprimée pour la Tox
  ExecuteSQL('UPDATE PIECE SET GP_SUPPRIME="X",GP_DATEMODIF="'+DateSupp+'",'+
             'GP_UTILISATEUR="'+V_PGI.User+'" WHERE '+WherePiece(CleDoc,ttdPiece,False));
  //Supprime les enregistrements des tables suivantes :
  ExecuteSQL('DELETE LIGNE WHERE '+WherePiece(CleDoc,ttdLigne,False));
  ExecuteSQL('DELETE PIECEADRESSE WHERE '+WherePiece(CleDoc,ttdPieceAdr,False));
  ExecuteSQL('DELETE PIEDBASE WHERE '+WherePiece(CleDoc,ttdPiedBase,False));
  ExecuteSQL('DELETE PIEDECHE WHERE '+WherePiece(CleDoc,ttdEche,False));
  ExecuteSQL('DELETE PIEDPORT WHERE '+WherePiece(CleDoc,ttdPorc,False));
//  ExecuteSQL('DELETE LIGNENOMEN WHERE '+WherePiece(CleDoc,ttdNomen,False));
//  ExecuteSQL('DELETE ACOMPTES WHERE '+WherePiece(CleDoc,ttdAcompte,False));
  //ExecuteSQL('DELETE LIGNELOT WHERE '+WherePiece(CleDoc,ttdLot,False));
  //ExecuteSQL('DELETE LIGNESERIE WHERE '+WherePiece(CleDoc,ttdSerie,False));
end;

// Procédure appelée par une transaction pour épurer toutes les Annonces de livraison sélectionnées
procedure TOF_MBOEPURATIONALF.SetAllEPURATION;
var DateSupp : string;
begin
  DateSupp := UsTime(NowH);

  //Supprime les enregistrements des tables suivantes :
  ExecuteSQL('DELETE LIGNE WHERE EXISTS (SELECT GP_NATUREPIECEG FROM PIECE '+StWhere+' AND LIEN(PIECE,LIGNE))');
  ExecuteSQL('DELETE PIECEADRESSE WHERE EXISTS (SELECT GP_NATUREPIECEG FROM PIECE '+StWhere+' AND LIEN(PIECE,PIECEADRESSE))');
  ExecuteSQL('DELETE PIEDBASE WHERE EXISTS (SELECT GP_NATUREPIECEG FROM PIECE '+StWhere+' AND LIEN(PIECE,PIEDBASE))');
  ExecuteSQL('DELETE PIEDECHE WHERE EXISTS (SELECT GP_NATUREPIECEG FROM PIECE '+StWhere+' AND LIEN(PIECE,PIEDECHE))');
  ExecuteSQL('DELETE PIEDPORT WHERE EXISTS (SELECT GP_NATUREPIECEG FROM PIECE '+StWhere+' AND LIEN(PIECE,PIEDPORT))');
//  ExecuteSQL('DELETE LIGNENOMEN WHERE EXISTS (SELECT GP_NATUREPIECEG FROM PIECE '+StWhere+' AND LIEN(PIECE,LIGNENOMEN))');
//  ExecuteSQL('DELETE ACOMPTES WHERE EXISTS (SELECT GP_NATUREPIECEG FROM PIECE '+StWhere+' AND LIEN(PIECE,ACOMPTES))');
  //ExecuteSQL('DELETE LIGNELOT WHERE EXISTS (SELECT GP_NATUREPIECEG FROM PIECE '+StWhere+' AND LIEN(PIECE,LIGNELOT))');
  //ExecuteSQL('DELETE LIGNESERIE WHERE EXISTS (SELECT GP_NATUREPIECEG FROM PIECE '+StWhere+' AND LIEN(PIECE,LIGNESERIE))');

  //Tag la pièce supprimée pour la Tox
  ExecuteSQL('UPDATE PIECE SET GP_SUPPRIME="X",GP_DATEMODIF="'+DateSupp+'",'+
             'GP_UTILISATEUR="'+V_PGI.User+'" '+StWhere);
end;



///////////////////////////////////////////////////////////////////

// procedure appellée par le bouton BOuvrir
procedure AGLBatchEPURATION_ALF(Parms : Array of Variant; Nb : Integer);
var F : TForm;
    TOTOF : TOF;
begin
F := TForm(Longint(Parms[0]));
if (F is TFmul) then TOTOF := TFMul(F).LaTOF else exit;
if (TOTOF is TOF_MBOEPURATIONALF) then TOF_MBOEPURATIONALF(TOTOF).BatchEPURATION_ALF;
end;

Initialization
  registerclasses ( [ TOF_MBOEPURATIONALF ] );
  RegisterAGLProc('BatchEPURATION_ALF', True, 0, AGLBatchEPURATION_ALF);
end.
