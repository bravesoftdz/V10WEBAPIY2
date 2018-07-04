unit SiscoVersPGI;

interface

uses
  SysUtils, Classes, Hctrls, Ent1, HEnt1,

{$IFDEF EAGLCLIENT}
{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}

{$IFNDEF EAGLSERVER}
  hmsgbox, FE_Main,
{$ELSE}
  UImpUtil,
  ULibCpContexte,
  UTOB,
  uWa,
{$ENDIF}
  HStatus,
  ImpFicU, TImpFic, RecupUtil,
  SISCO, RecordCom, UtilTrans;


Function TransfertSiscoVersPgi(StFichier : String ; RemplaceLeFichier : Boolean ; SansAux : Boolean=TRUE; NewRep : String = '' ; InfoImp : PtTInfoImport = Nil; Starg : string='') : string ;

implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  CPTypeCons,
  CPProcGen,
  {$ENDIF MODENT1}
  ParamSoc;

Procedure AfficheMessageServer (InfoMessage,Element : string);
begin
{$IFNDEF EAGLSERVER}
                         PGIBox (InfoMessage, Element);
{$ELSE}
                         ddWriteLN(InfoMessage);
                         cWA.MessagesAuClient('COMSX.IMPORT','', InfoMessage);
{$ENDIF}

end;

Function TraiteEnrVersPGI(Var NewFichier : TextFile ; ListePiece : HTStringList ;
                   Var St,St1 : String ; Var IdentPiece : TIdentPiece ;
                   Var StSISCO : TStSISCO ; Var Info : TInfoGeneSISCO ;
                   Var CloseOpenAFaire : Boolean ;
                   Var ExisteExerc : boolean;
                   InfoImp : PtTInfoImport;
                   SansAux : Boolean; StFichier, Starg : string; var Decimal : string) : Integer ;
Var
    StSup          : String ;
    OkPiece        : Integer ;
    i              : Integer ;
    PieceAvecMvt,RuptOrigineSaisie,ChronoEclateIncremente : Boolean ;
    Ligne          : string;
    CodeDevise,Eur : string;
    Q1             : TQuery;
    Code           : string;
    D1,D2          : TDateTime;
    lg,lg2         : integer;
    lgFichier      : integer;
//    ExisteExerc    : boolean;
{$IFNDEF EAGLSERVER}
    RetourConf      : string;
{$ENDIF}
BEGIN
Result:=0 ; StSup:='' ; CloseOpenAFaire:=FALSE ;
If Copy(St,1,2)='00' Then
BEGIN
         ExisteExerc := ExisteSQl ( 'SELECT EX_EXERCICE from EXERCICE');
         if not ExisteExerc then
         begin
            SetParamSoc('SO_TENUEEURO', TRUE);
         end;
         Decimal := Copy(St, 23, 1);  // fiche 10577
         Info.ExoSISCO.Deb:=Str6ToDate(Copy(St,8,6),50) ;
         Info.ExoSISCO.Fin:=Str6ToDate(Copy(St,14,6),50) ;
         D1 := Format_Date_HAL(FormatDateTime(Traduitdateformat('ddmmyyyy'), Info.ExoSISCO.Deb));
         D2 := Format_Date_HAL(FormatDateTime(Traduitdateformat('ddmmyyyy'), Info.ExoSISCO.Fin));
         Q1 := OpenSQL('SELECT EX_EXERCICE from EXERCICE Where EX_DATEDEBUT="' + UsDateTime(D1) + '" AND EX_DATEFIN="' + UsDateTime(D2)+'"', TRUE);
         if Q1.EOF then
             Code := '001'
         else
             Code := Q1.FindField ('EX_EXERCICE').asstring;
         Ferme (Q1);
{$IFNDEF NOVH}
         VH^.Cpta[fbGene].Lg := StrToInt (Copy(St,21,2));
         VH^.Cpta[fbAux].Lg := StrToInt (Copy(St,21,2));
{$ELSE}
         lgFichier := StrToInt (Copy(St,21,2));
{$ENDIF}
         if Copy(St,36,3) = 'FRF' then   // fiche 10294
         begin
                        AfficheMessageServer  ('Import impossible, l''exercice est en Franc ' , 'Import Sisco II');
                        Result := -1;
                        exit;
         end;
         if Copy(St,20,1) = 'E' then    // ajout me 10-01-2003
         begin
{$IFNDEF NOVH}
               VH^.RecupComSx := TRUE; //pour récuperer le lettrage
{$ENDIF}

               Q1 := OpenSQL('SELECT EX_EXERCICE,EX_ETATCPTA,EX_LIBELLE from EXERCICE Where EX_ETATCPTA="OUV"  and EX_EXERCICE <"' + Code + '"', TRUE);
               if not Q1.EOF then
               begin
                        AfficheMessageServer  ('Import impossible, l''exercice ne correspond pas à l''exercice en cours du dossier comptable ' + #10#13 +'('+ Q1.FindField ('EX_LIBELLE').asstring +')', 'Import Sisco II');
                         Ferme (Q1); Result := -1;
                        exit;
               end;
               Ferme (Q1);
               Q1 := OpenSQL('SELECT EX_EXERCICE from EXERCICE Where EX_ETATCPTA="CDE"  and EX_EXERCICE ="' + Code + '"', TRUE);
               if not Q1.EOF then
               begin
                        AfficheMessageServer  ('Exercice est clos', 'Import Sisco II');
                         Ferme (Q1); Result := -1;
                        exit;
               end;
               Ferme (Q1);
               // Exemple
               // SELECT E_DATEPAQUETMAX from ECRITURE Where E_ETATLETTRAGE="TL" and E_LETTRAGE <> "" and
               // and E_EXERCICE="004" and E_DATEPAQUETMAX > "12/31/2003"

               Q1 := OpenSQL('SELECT E_DATEPAQUETMAX from ECRITURE Where E_ETATLETTRAGE="TL" and E_LETTRAGE <> "" and E_EXERCICE="'+Code +'"'+' and E_DATEPAQUETMAX >"' + UsDateTime(D2) + '"', TRUE);
               if not Q1.EOF then
               begin
                        AfficheMessageServer  ('Des écritures sont lettrées entre N et N+1. Traitement impossible.', 'Import Sisco II');
                         Ferme (Q1); Result := -1;
                        exit;
               end;
               Ferme (Q1);
{$IFNDEF EAGLSERVER}
         if Starg = '' then
         begin
               RetourConf := AglLanceFiche('CP','IMPORTCOMCONF','', '',
               'Origine du fichier : SI; Nature : EXE'
                        +'; Fichier : '+StFichier+';'+Code)  ;
               if RetourConf = '-' then
               begin
                    Result := -1;
                    exit;
               end
               else
                  ExisteExerc := ExisteSQl ( 'SELECT EX_EXERCICE from EXERCICE');
         end
         else
         begin
               ExecuteSQL('DELETE FROM ECRITURE Where E_EXERCICE="'+Code + '"');
               ExecuteSQL('DELETE FROM ANALYTIQ Where Y_EXERCICE="'+Code + '"');
         end;
{$ENDIF}

         end;

  if Copy(St,20,1) = 'E' then
     Ecritureentete(NewFichier, 'DOS', 'STD','SIS')
  else
      if Copy(St,20,1) = 'B' then
         Ecritureentete(NewFichier, 'BAL', 'STD','SIS')
         else
             if Copy(St,20,1) = 'J' then
                Ecritureentete(NewFichier, 'JRL', 'STD','SIS');
{$IFNDEF NOVH}
         lgFichier := VH^.Cpta[fbGene].Lg;
{$ELSE}
         lgFichier := StrToInt (Copy(St,21,2));
{$ENDIF}

      lg :=  GetParamSocSecur('SO_LGCPTEGEN', lgFichier);
      lg2 := GetParamSocSecur('SO_LGCPTEAUX', lgFichier);

      if (Copy(St,20,1) = 'E') and ((lg = 0) or (lg2 = 0)) then
      begin
            SetParamsoc('SO_LGCPTEGEN', lgFichier);
            SetParamsoc('SO_LGCPTEAUX', lgFichier);
            lg := lgFichier;
            lg2 := lgFichier;
{$IFDEF NOVH}
            TCPContexte.GetCurrent.InfoCpta.ChargeLgDossier ;
{$ENDIF}
      end;
      if (lgFichier = lg)  and (lgFichier = lg2) then
         Info.Troncaux := FALSE
      else
         Info.Troncaux := TRUE;

(* Plus qe question car traitement est séparé des écran
    if not SansAux then
    begin
      Okrep := PGiAskCancel('Voulez-vous importer les données en enlevant le premier caractère des comptes auxiliaires ?','Import Sisco II') ;
      if Okrep = mrCancel then begin Result := -1; exit; end;
      if Okrep = mrYes then Info.Troncaux := TRUE
      else
      Info.Troncaux := FALSE;
    end
    else
    Info.Troncaux := FALSE;
*)
    RecupSociete(0,St,Info,InfoImp)  ;

    if not Info.Troncaux then
    begin
      if (GetInfoCpta(fbGene).Lg <> lg) and (Lg <> 0) or
               (GetInfoCpta(fbAux).Lg <> lg2) and (Lg2 <> 0) then
      begin
           AfficheMessageServer ('La longueur des comptes est différente', 'Impossible de remonter le fichier');
           Result := -1;
           SetParamSoc('SO_LGCPTEGEN',lg) ;
           SetParamSoc('SO_LGCPTEAUX',lg2) ;
{$IFNDEF NOVH}
           VH^.Cpta[fbGene].Lg:=lg ;
           VH^.Cpta[fbAux].Lg:=lg2 ;
{$ELSE}
           TCPContexte.GetCurrent.InfoCpta.ChargeLgDossier ;
{$ENDIF}
           // car recupsociete créee l'exercice
           if not ExisteExerc then ExecuteSQL('DELETE FROM EXERCICE');
           Exit;
      end;
    end
    else
    begin
      if (GetInfoCpta(fbGene).Lg-1 <> lg) and (lg <> 0) or
               (GetInfoCpta(fbAux).Lg-1 <> lg2) and (lg2 <> 0) then
      begin
           AfficheMessageServer ('Import du fichier impossible. La  longueur des comptes est incohérente avec le paramétrage du dossier.', 'Import Sisco II');
           Result := -1;
           SetParamSoc('SO_LGCPTEGEN',lg) ;
           SetParamSoc('SO_LGCPTEAUX',lg2) ;
{$IFNDEF NOVH}
           VH^.Cpta[fbGene].Lg:=lg ;
           VH^.Cpta[fbAux].Lg:=lg2 ;
{$ELSE}
           TCPContexte.GetCurrent.InfoCpta.ChargeLgDossier ;
{$ENDIF}
           // car recupsociete créee l'exercice
           if not ExisteExerc then ExecuteSQL('DELETE FROM EXERCICE');
           Exit;
      end;
      if (GetInfoCpta(fbGene).Lg <> lg) and (lg <> 0) or
               (GetInfoCpta(fbAux).Lg <> lg2) and (lg2 <> 0) then
      begin
           SetParamSoc('SO_LGCPTEGEN',lg) ;
           SetParamSoc('SO_LGCPTEAUX',lg2) ;
{$IFNDEF NOVH}
           VH^.Cpta[fbGene].Lg:=lg ;
           VH^.Cpta[fbAux].Lg:=lg2 ;
{$ELSE}
           TCPContexte.GetCurrent.InfoCpta.ChargeLgDossier ;
{$ENDIF}
      end;
    end;

    If not ExisteExerc then
    Code := '001'
    else
    begin
         D1 := Format_Date_HAL(FormatDateTime(Traduitdateformat('ddmmyyyy'), Info.ExoSISCO.Deb));
         D2 := Format_Date_HAL(FormatDateTime(Traduitdateformat('ddmmyyyy'), Info.ExoSISCO.Fin));
         Q1 := OpenSQL('SELECT EX_EXERCICE from EXERCICE Where EX_DATEDEBUT="' + UsDateTime(D1) + '" AND EX_DATEFIN="' + UsDateTime(D2)+'"', TRUE);
         if Q1.EOF then
         begin
           Ferme (Q1);
           if not VerifCoherenceExo ( Info.ExoSISCO.Deb, Info.ExoSISCO.Fin ) then
           begin
                        AfficheMessageServer  ('Les dates d''exercice du fichier sont incohérentes par rapport au dossier. Traitement impossible.', 'Import Sisco II');
                        Result := -1;
                        exit;
           end;
           Q1 := OpenSQl ('select max(ex_exercice) from exercice', TRUE);
           if not Q1.EOF then
           FormatFloat('000',StrToInt(Q1.Fields[0].asstring)+1)
         end
         else
          Code := Q1.FindField ('EX_EXERCICE').asstring;
         Ferme (Q1);
    end;
//  Ligne := Format(Formatexercice, ['001',
  Ligne := Format(Formatexercice, [Code,
  FormatDateTime(Traduitdateformat('ddmmyyyy'), Info.ExoSISCO.Deb),
  FormatDateTime(Traduitdateformat('ddmmyyyy'), Info.ExoSISCO.Fin),
  'OUV', 'OUV',
  'Exo de '+ FormatDateTime(Traduitdateformat('dd/mm/yyyy'), Info.ExoSISCO.Deb) + ' au ' + FormatDateTime(Traduitdateformat('dd/mm/yyyy'), Info.ExoSISCO.Fin),'H' ]);
  writeln(NewFichier, Ligne);
{$IFNDEF NOVH}
  VH^.EnCours.Deb:=Info.ExoSISCO.Deb ;
  VH^.EnCours.Fin:=Info.ExoSISCO.Fin ;
  VH^.EnCours.Code:=Code ;
{$ENDIF}

END Else
If Copy(St,1,2)='01' Then RecupSociete(1,St,Info,InfoImp) Else
If Copy(St,1,2)='02' Then RecupSociete(2,St,Info,InfoImp)
Else
If Copy(St,1,2)='03' Then
BEGIN
  RecupSociete(3,St,Info,InfoImp);
  Ligne := Format(Formatgeneraux1, [Info.Nom,
    Info.Adr1, Info.Adr2,Info.Adr3,
      copy(Info.Adr3, 0, 5) , copy (Info.Adr3, 6, 26),
      '', Info.Tel,
      '', '', '','','', '', Info.Siret,'',Info.APE, '','']);
  writeln(NewFichier, Ligne);
  if Info.Troncaux then
  begin
       lg := StrToint (Info.lgcpt)-1;
       if (lg >= 6) and (lg <= 10) then
       Info.lgcpt := inttoStr(lg);
{$IFNDEF NOVH}
       VH^.Cpta[fbGene].Lg:= lg;
       VH^.Cpta[fbAux].Lg:= lg;
{$ELSE}
         SetParamsoc('SO_LGCPTEGEN', lg);
         SetParamsoc('SO_LGCPTEAUX', lg);
{$ENDIF}
  end;

If not ExisteExerc then
  Ligne := Format(Formatgeneraux2, [Info.lgcpt,
    '0', Info.lgcpt,
// voir les sections analytiques
    '0', '6', '0', '6', '0', '6', '0', '6', '0', '6', '0',
    copy ('4720000000', 0, StrToint(Info.lgcpt)), copy(St, 57,1)+copy ('9999999999', 0,StrToint(Info.lgcpt)),
    copy(St, 56,1)+copy ('0000000000',0, StrToint(Info.lgcpt)), copy ('4219999999', 0,StrToint(Info.lgcpt)),
    copy ('4710000000',0,StrToint(Info.lgcpt)), '', '', '','','', '', '-', '-', '-', '-'])
else   // fiche 10514
  Ligne := Format(Formatgeneraux2, [Info.lgcpt,
    '0', Info.lgcpt,
// voir les sections analytiques
    '0',  IntToStr(GetInfoCpta(fbaxe1).Lg), '0',  IntToStr(GetInfoCpta(fbaxe2).Lg), '0',  IntToStr(GetInfoCpta(fbaxe3).Lg), '0',  IntToStr(GetInfoCpta(fbaxe4).Lg), '0',  IntToStr(GetInfoCpta(fbaxe5).Lg), '0',
    copy ('4720000000', 0, StrToint(Info.lgcpt)), copy(St, 57,1)+copy ('9999999999', 0,StrToint(Info.lgcpt)),
    copy(St, 56,1)+copy ('0000000000',0, StrToint(Info.lgcpt)), copy ('4219999999', 0,StrToint(Info.lgcpt)),
    copy ('4710000000',0,StrToint(Info.lgcpt)), '', '', '','','', '', '-', '-', '-', '-']);


  writeln(NewFichier, Ligne);
  If Info.EnEuro Then
  begin
     CodeDevise := 'EUR';
     Eur := 'X';
  end else
  begin
     CodeDevise := 'FRF';
     Eur := '-';
  end;
  if Decimal = '' then Decimal := '2';  // fiche 10577
  Ligne := Format(Formatgeneraux5, [CodeDevise,
    Decimal, Eur, '',copy('6688000000', 0, StrToint(Info.lgcpt)),
      copy ('7688000000',0, StrToint(Info.lgcpt)), '6,55957','','', '', '','', Info.Numplan, '-', GetParamSocSecur('SO_CROISAXE', False), '-']);
  writeln(NewFichier, Ligne);
END
else
If Copy(St,1,2)='04' Then
begin
     if (not ExisteExerc ) and (Copy(St,42,1) = '1') then // Exercice clos
     begin
          AfficheMessageServer  ('L''import d''un exercice clos est impossible.', 'Import Sisco II');
          Result := -1;
          exit;
     end;
     RecupSociete(4,St,Info,InfoImp);

end
else
If (Copy(St,1,2)='08') and ((Copy(St,33,1) = 'C') or (Copy(St,33,1) = 'F')) Then
begin
     RecupSociete(8,St,Info,InfoImp);
     ChargeCptSISCO(Info.FourchetteSISCO,'') ;
     For i:=0 To 50 Do
     BEGIN
       if Info.FourchetteSISCO[i].Cpt='' Then break ;
       if (Info.FourchetteSISCO[i].Cpt >= Info.FourchetteSISCO[i].Aux1)  and (Info.FourchetteSISCO[i].Aux2 >= Info.FourchetteSISCO[i].Cpt) then
       begin
          AfficheMessageServer ('Attention les tranches des comptes auxiliaires ne sont pas conformes avec le collectif associé.#10#13Enregistrement 08 incorrect.', 'Import Sisco II');
          Result := -1; Exit;
       end;
     END ;
end Else
If Copy(St,1,2)='50' Then RecupSociete(50,St,Info,InfoImp) Else
If Info.InitPCL Then
  If Copy(St,1,2)='05' Then BEGIN If RecupJalSISCO(St,St1,Info,InfoImp) Then WriteLn(NewFichier,St1) ;END Else If Info.InitPCL Then
    BEGIN
    Case St[1] Of
      'C' : BEGIN
      RecupCptSISCO(St,St1,StSup,Info,InfoImp, 'CAE') ; WriteLn(NewFichier,St1) ; If GetRecupPCL And (StSup<>'') Then WriteLn(NewFichier,StSup) ; END ;    // Généraux & Auxiliaires
      'S' : BEGIN RecupSectionSISCO(St,St1,Info) ; WriteLn(NewFichier,St1) ; END ; // Section
      'M' : BEGIN
            EcritPiece(NewFichier,ListePiece,IdentPiece,Info,PieceAvecMvt,InfoImp) ;
            OkPiece:=FaitPiece(St, ListePiece, IdentPiece, StSISCO, Info, InfoImp, FALSE) ;
{$IFDEF TTW}
  AfficheMessageServer('OKPiece :'+IntToStr(OkPiece), '');
{$ENDIF}

            If GetRecupPCL And (OkPiece<>0) Then BEGIN Info.PbEnr[0]:=TRUE ; END ;
            IdentPiece.NumChronoEclate:=0 ; ChronoEclateIncremente:=FALSE ;
            END ; // Mois
      'J' : BEGIN
            EcritPiece(NewFichier,ListePiece,IdentPiece,Info,PieceAvecMvt,InfoImp) ;
            FaitPiece(St, ListePiece, IdentPiece, StSISCO, Info, nil, FALSE) ;
            IdentPiece.NumChronoEclate:=0 ; ChronoEclateIncremente:=FALSE ;
            END ; // Journal
      'F' : BEGIN
            ChronoEclateIncremente:=FALSE ;
            If GetRecupPCL Then
              BEGIN
              // PieceEquilibree(IdentPiece) ;
              IdentPiece.FolioSISCO:='Folio '+Trim(Copy(St,2,3)) ;
              // ajout me 27-05-2002 pour le découpage
              IdentPiece.NoFolioSISCO := 1;
              END Else
              BEGIN
              IdentPiece.FolioSISCO:='Folio '+Trim(Copy(St,2,3)) ; IdentPiece.NoFolioSISCO:=StrToInt(Trim(Copy(St,2,3))) ;
              END ;
            END ;// Rien à faire
      'E' : BEGIN
            RuptOrigineSaisie:=RuptureOrigineSaisie(St,IdentPiece,ChronoEclateIncremente) ;
            If Not RuptOrigineSaisie Then ChronoEclateIncremente:=FALSE ;
            If GetRecupPCL Then
              // PieceEquilibree(IdentPiece)
            Else
            BEGIN
              If RuptureJour(St,StSisco) Or PieceEquilibree(IdentPiece) Or RuptureOrigineSaisie(St,IdentPiece,ChronoEclateIncremente) Then EcritPiece(NewFichier,ListePiece,IdentPiece,Info,PieceAvecMvt,InfoImp) ;
            END ;
{$IFDEF TTW}
  AfficheMessageServer('AVT FaitPiece', '');
{$ENDIF}

            FaitPiece(St, ListePiece, IdentPiece, StSISCO, Info, InfoImp, RuptOrigineSaisie) ;
{$IFDEF TTW}
  AfficheMessageServer('APRES FaitPiece', '');
{$ENDIF}
            END ; // Ecriture
      'p' : ; // Non traité pour l'instant
      'v' : BEGIN FaitPiece(St, ListePiece, IdentPiece, StSISCO, Info, InfoImp, FALSE) ; END ;
      'A' : ; // Non traité pour l'instant
      'B' : ; // Non traité pour l'instant
      END ;
    END ;
END ;


Function TransfertSiscoVersPgi(StFichier : String ; RemplaceLeFichier : Boolean ; SansAux : Boolean=TRUE; NewRep : String = '' ; InfoImp : PtTInfoImport = Nil; Starg : string='') : string ;
Var Fichier,NewFichier : TextFile ;
    St,St1,StNewFichier : String ;
    ListePiece : HTStringList ;
    IdentPiece : TIdentPiece ;
    StSISCO : TStSISCO ;
    InfoGeneSISCO : TInfoGeneSISCO ;
    Pb : Boolean;
    {$IFNDEF EAGLSERVER}
    FichierOk : Boolean ;
    CompteurFichier : Integer ;
    {$ENDIF}
    ProfilPCL : String ;
    OkOk : Boolean ;
    PieceAvecMvt : Boolean ;
    CloseOpenAFaire : Boolean ;
    NewFich         : string;
    rr              : integer;
    ExisteExerc     : boolean;
    Decim           : string;
BEGIN
rr :=0; Result:='' ; Pb:=FALSE ;
{$IFNDEF NOVH}
VH^.RecupPCL:=TRUE ;
{$ENDIF}
InitMove(1000,TraduireMemoire('Chargement du fichier en cours...')) ;
AssignFile(Fichier,StFichier) ; Reset(Fichier) ;
{$IFDEF EAGLSERVER}
StNewFichier := ChangeFileExt(StFichier,'.TRA');
NewFich := StNewFichier;
{$ELSE}
StNewFichier := FileTemp('.PNM') ;
{$ENDIF}
AssignFile(NewFichier, StNewFichier) ; Rewrite(NewFichier) ;
ReadLn(Fichier,St) ; // Début du fichier SISCO
ListePiece:=HTStringList.Create ;
ListePiece.Sorted:=FALSE ; //ListePiece.Duplicates:=dupAccept ;
Fillchar(IdentPiece,SizeOf(IdentPiece),#0) ; IdentPiece.NumP:=1 ;

// Ecritureentete(NewFichier, 'DOS', 'STD','SIS');

Fillchar(StSISCO,SizeOf(StSISCO),#0) ;
Fillchar(InfoGeneSISCO,SizeOf(InfoGeneSISCO),#0) ;
ProfilPCL:='' ;
If InfoImp<>NIL Then If InfoIMP^.ProfilPCL<>'' Then ProfilPCL:=InfoIMP^.ProfilPCL ;
InfoGeneSISCO.InitPcl := TRUE;
ChargeCptSISCO(InfoGeneSISCO.FourchetteSISCO,ProfilPCL) ;
ChargeFourchetteCompte('SIS',InfoGeneSISCO.FourchetteImport) ;
ChargeCharRemplace(InfoGeneSISCO) ;
InfoGeneSISCO.LJalLu:=HTStringList.Create ;
{$IFNDEF EAGLSERVER}
CompteurFichier:=-1 ;
{$ENDIF}
InitRequete(InfoGeneSISCO.QFiche[3],3) ;
While (Not EOF(Fichier)) And (Not Pb)  do
  BEGIN
  MoveCur(FALSE) ;
  ReadLn(Fichier, St) ;
  OkOk:=TRUE ;
  if St = '' then break;
  If OkOk Then rr:=TraiteEnrVersPGI(NewFichier,ListePiece,St,St1,IdentPiece,StSISCO,InfoGeneSISCO,CloseOpenAFaire,ExisteExerc,InfoImp,
SansAux,StFichier, Starg, Decim) ;
 If InfoGeneSISCO.PbEnr[0] Then Break ;
  If rr<>0 Then Pb:=TRUE ;
  END ;
  if Pb then begin
     InfoImp^.PbEnr[0] := TRUE;
      CloseFile(Fichier) ; CloseFile(NewFichier) ;
      Ferme(InfoGeneSISCO.QFiche[3]) ;
      VideStringList(ListePiece) ; ListePiece.Free ;
      VideStringList(InfoGeneSISCO.LJalLu) ; InfoGeneSISCO.LJalLu.Free ;
     exit;
  end;
EcritPiece(NewFichier,ListePiece,IdentPiece,InfoGeneSISCO,PieceAvecMvt,InfoImp) ;
FiniMove ;
Flush(NewFichier) ;
CloseFile(Fichier) ; CloseFile(NewFichier) ;
Ferme(InfoGeneSISCO.QFiche[3]) ;
VideStringList(ListePiece) ; ListePiece.Free ;
VideStringList(InfoGeneSISCO.LJalLu) ; InfoGeneSISCO.LJalLu.Free ;

{$IFNDEF EAGLSERVER}
If rr=0 Then NewFich := CloseEtRenameNewFile(RemplaceLeFichier,StFichier,StNewFichier,NewRep,CompteurFichier) ;
InitMove(500,'') ;
Repeat
  MoveCur(FALSE) ;
  AssignFile(Fichier,StFichier) ;
  {$I-} Reset (Fichier) ; {$I+}
  FichierOk:=ioresult=0 ;
  If FichierOk Then Close(Fichier) ;
Until FichierOk ;
FiniMove ;
{$ENDIF}

If GetRecupPCL And (InfoImp<>Nil) Then
  BEGIN
  InfoImp^.AuMoinsUnClient:=InfoGeneSISCO.AuMoinsUnClient ;
  InfoImp^.AuMoinsUnFournisseur:=InfoGeneSISCO.AuMoinsUnFournisseur ;
  InfoImp^.OkFouCli:=TRUE ;
  InfoImp^.OkFouFou:=TRUE ;
  If (InfoImp^.AuMoinsUnFournisseur) And (Not FourchetteSISCOExiste(0,InfoGeneSISCO.FourchetteSISCO)) Then InfoImp^.OkFouFou:=FALSE ;
  If (InfoImp^.AuMoinsUnClient) And (Not FourchetteSISCOExiste(1,InfoGeneSISCO.FourchetteSISCO)) Then InfoImp^.OkFouCli:=FALSE ;
  InfoImp^.PbEnr:=InfoGeneSISCO.PbEnr ;
  END ;
  Result := NewFich;
END ;


end.
