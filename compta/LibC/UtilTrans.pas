unit UtilTrans;

interface
uses
     classes, SysUtils, windows, HEnt1, ComCtrls, StdCtrls,ExtCtrls, hmsgbox, controls,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HCtrls,Graphics,ed_tools, UTOB, forms,
{$IFDEF EAGLSERVER}
     esession, ULibCpContexte, FileCtrl, galSystem,eIsapi,
{$ENDIF}
     paramsoc;

Function QuelMontant(Q : TQuery ; P : String ; b : Byte ; Decim : Integer ; Var Sens : String) : String ;
Function AlignDroite(St : String ; l : Integer) : String ;
function Format_Date_HAL(Dat: string): TDateTime;
Function AfficheListeCom(Chaine: string; Listecom : TListBox; InitProgess : Boolean=FALSE) : Boolean;
FUNCTION STSTRFPOINT(St : String) : string ;
Function TransIsValidDate ( dte : string) : Boolean ;
procedure EcrireDansfichierListeCom (Fichier : string; Listecom : TListBox);
Function RendCommandeJournal(ex,ts : string) : string;
Function RendCommandeJournalBud(ex,ts : string) : string;
Function RendCommandeExo(ex,ts : string) : string;
Function RendCommandeComboMulti(ex,ts,ch : string) : string;
FUNCTION AGauche (st : string; l : Integer ; C : Char) : string ;
FUNCTION ADroite (st : string; l : Integer ; C : Char) : string ;
Procedure ControleCompte (var Compte : string);
Procedure GrisechampEdit(champ: THEdit; En : Boolean) ;
Procedure GrisechampCombo(champ: TComboBox; En : Boolean) ;
Procedure GrisechampCritEdit(champ: THCritMaskEdit; En : Boolean) ;
Procedure InitProgressbar(Titre : string='Execution Export ');
Function AfficheProgressbar( NN : string) : Boolean;
Procedure FiniProgressbar;
Function OkLettrage (Where : string; Listb : TListBox; TOBLET : TOB=nil) : Boolean;
procedure InitDossierEnNombre (Var GrandNbGene, GrandNbAux : Boolean ; WhereGene : string=''; WhereAux : string=''); // AJOUT ME 14-01-2005
Function GetPeriodeJour ( LaDate : TDateTime ) : integer ;
procedure UpdateComLettrage(What : Byte; Condition : string='');
procedure EpureChar (Buf: string; var Resultat: string);
function SupprimeFichier(FileN: string; Silence: Boolean = False): Boolean;
function PGUpperCase(const S: string; withAccent: boolean = true): string;
function PGSupprimeaccent(Ch: Char): Char;
procedure PaysISOLib(Pays: string; var CodeIso, Libelle: string);

{$IFDEF EAGLSERVER}
Function RenseigneWrootPath (WhatAppli : string) : string;
procedure LanceCreatDir (Appli : string);
Procedure LanceRemoveDirectory (Appli : string);
{$ENDIF}

Const
DebugS1 : Boolean = FALSE;
MaxNombre = 2000; // AJOUT ME 14-01-2005
{$IFDEF EAGLCLIENT}
NameAppli = 'ECOMSX';
{$ELSE}
NameAppli = 'COMSX';
{$ENDIF}


implementation

Function QuelMontant(Q : TQuery ; P : String ; b : Byte ; Decim : Integer ; Var Sens : String) : String ;
Var MD,MC,Montant : Double ;
    Sup : String ;
    LeSens : String ;
BEGIN
Result:='' ; Sup:='' ; Montant:=0 ;
Case b Of 1 : Sup:='DEV' ; (*2 : ajout me 18-07-2003 Sup:='EURO' ;*) END ;
MD:=Q.Findfield(P+'_DEBIT'+Sup).AsFloat ; MC:=Q.Findfield(P+'_CREDIT'+Sup).AsFloat ;
MD:=Arrondi(MD,Decim) ; MC:=Arrondi(MC,Decim) ;
if ((MD<0) or (MD>0)) and (MC=0) then BEGIN Montant:=MD ; LeSens:='D' ; END else
 if ((MC<0) or (MC>0)) and (MD=0) then BEGIN Montant:=MC ; LeSens:='C' ; END ;
If Sens='' Then Sens:=LeSens ;
Result:=StrfMontant(Montant,20,Decim,'',False) ;
END ;

Function AlignDroite(St : String ; l : Integer) : String ;
var St1 : String ;
BEGIN
St1:=Trim(st) ;
While Length(St1)<l Do St1:=' '+St1 ;
AlignDroite:=St1 ;
END ;

function Format_Date_HAL(Dat: string): TDateTime;
var
  An: string;
  Y, M, J: Word;
  StD: string;
begin
  Result := iDate1900;
  if Trim(Dat) = '' then
    Exit;
  An := Copy(Dat, 5, 4);
  Y := Round(Valeur(An));
  M := Round(Valeur(Copy(Dat, 3, 2)));
  J := Round(Valeur(Copy(Dat, 1, 2)));
  if (J = 0) or (M = 0) or (Y = 0) then
    Exit;
  if j in [1..31] = FALSE then
    Exit;
  if m in [1..12] = FALSE then
    Exit;
  StD := FormatFloat('00', J) + '/' + FormatFloat('00', M) + '/' + An;
  if not IsValidDate(StD) then
  begin
    if m <> 2 then
      J := 30
    else
      J := 28;
  end;
  Result := EncodeDate(Y, M, J);
end;

Procedure InitProgressbar(Titre : string='Execution Export ');
begin
       InitMoveProgressForm (nil, Titre, 'Traitement en cours ...', 1000, TRUE, TRUE) ;
end;

Function AfficheProgressbar( NN : string) : Boolean;
begin
        Result := TRUE;
        if not (MoveCurProgressForm (NN)) then Result := FALSE;
end;

Procedure FiniProgressbar;
begin
           FiniMoveProgressForm;
end;

Function AfficheListeCom(Chaine: string; Listecom : TListBox; InitProgess : Boolean=FALSE) : Boolean;
var
Ch : string;
begin
  Result := TRUE;
  if ListeCom = nil then exit;
  Ch := TraduireMemoire(Chaine);
  Listecom.Items.add (Ch);
  Listecom.ItemIndex := Listecom.Items.Count-1;
  {$IFNDEF EAGLSERVER}
  if InitProgess then Result := AfficheProgressbar(Ch);
  {$ENDIF}
end;

procedure EcrireDansfichierListeCom (Fichier : string; Listecom : TListBox);
var
  F: TextFile;
  i: Integer;
  Ligne : string;
begin
{$I-}
  AssignFile(F, Fichier);
  if (FileExists(Fichier)) then
     Append(F)
  else
     Rewrite(F) ;
  writeln(F, 'Compte rendu d''import/export');
  for i := 0 to (Listecom.Items.Count - 1) do
  begin
     Ligne := Listecom.Items[i];
     writeln(F, Ligne);
  end;
  CloseFile(F);
{$I+}
end;

FUNCTION STSTRFPOINT(St : String) : string ;
var
    i : integer ;
BEGIN
//St:=FloatToStr(Tot) ;
i:=Pos(',',St) ; if i>0 then St[i]:='.' ;
Result:=St ;
END ;

Function TransIsValidDate ( dte : string) : Boolean ;
var
aa, mm, jj : word;
dd         : TDatetime;
begin
Result := IsValidDate( dte);
if Result then
begin
dd := StrToDate (dte);
DecodeDate (dd, aa, mm, jj);
if aa < 1900 then Result := FALSE
end;

end;


Function RendCommandeJournal(ex,ts : string) : string;
var
St       : string;
x        : string;
begin
St := '';
x := ReadTokenSt (ts);
if x <> '' then  St := ex+'_JOURNAL="' +x+ '"';
while x <> '' do
  begin
       x := ReadTokenSt (ts);
       if x <> '' then
       St := St+ ' OR '+ex+'_JOURNAL="' +x+ '"';
  end;
  Result := St;
end;

Function RendCommandeJournalBud(ex,ts : string) : string;
var
St       : string;
x        : string;
begin
St := '';
x := ReadTokenSt (ts);
if x <> '' then  St := ex+'_BUDJAL="' +x+ '"';
while x <> '' do
  begin
       x := ReadTokenSt (ts);
       if x <> '' then
       St := St+ ' OR '+ex+'_BUDJAL="' +x+ '"';
  end;
  Result := St;
end;

Function RendCommandeExo(ex,ts : string) : string;
var
St       : string;
x        : string;
begin
St := '';
x := ReadTokenSt (ts);
if x <> '' then  St := '('+ ex+'_EXERCICE="' +x+ '"';
while x <> '' do
begin
       x := ReadTokenSt (ts);
       if x <> '' then
       St := St+ ' OR '+ex+'_EXERCICE="' +x+ '"';
end;
if St <> '' then
     Result := St+')';
end;

Function RendCommandeComboMulti(ex,ts,ch : string) : string;
var
St       : string;
x        : string;
begin
St := '';
x := ReadTokenSt (ts);
if x <> '' then  St := '('+ ex+'_'+ch+'="' +x+ '"';
while x <> '' do
begin
       x := ReadTokenSt (ts);
       if x <> '' then
       St := St+ ' OR '+ex+'_'+ch+'="' +x+ '"';
end;
if St <> '' then
     Result := St+')';
end;

FUNCTION AGauche (st : string; l : Integer ; C : Char) : string ;
var St1 : String ;
BEGIN
st1:=Trim(st) ;
if ((l>0) and (l<Length(St1))) then St1:=Copy(St1,1,l) ;
While Length(St1)<l Do St1:=St1+C ;
Result:=st1 ;
END ;

FUNCTION ADroite (st : string; l : Integer ; C : Char) : string ;
var St1 : String ;
BEGIN
st1:=Trim(st) ;
if ((l>0) and (l<Length(St1))) then St1:=Copy(St1,1,l) ;
While Length(St1)<l Do St1:=C+St1 ;
Result:=st1 ;
END ;

Procedure ControleCompte (var Compte : string) ;
var
i  :integer;
Ok : Boolean;
St1: string;
begin
 for  i:=1  to length(Compte)-1 do
 begin
   St1:=Copy(Compte,i,1);
   Ok := ((St1>='0') and (St1<='9'))  or
   ((St1>='A') and (St1<='Z')) or
   ((St1>='a') and (St1<='z')) ;
   if not Ok then
      Compte[i] := '0';
 end;
end;

Procedure GrisechampEdit(champ: THEdit; En : Boolean) ;
begin
if (champ <> NIL) then
  begin
          champ.tabstop:=En;
          Champ.Enabled :=En;
          Champ.ReadOnly := not En;
          if En then
              champ.color :=ClWindow
          else
              champ.color :=ClBtnFace;
  end;
end;

Procedure GrisechampCombo(champ: TComboBox; En : Boolean) ;
begin
if (champ <> NIL) then
  begin
       champ.tabstop:=En;
       Champ.Enabled:=En;
       if En then
              champ.color :=ClWindow
       else
              champ.color :=ClBtnFace;
  end;
end;

Procedure GrisechampCritEdit(champ: THCritMaskEdit; En : Boolean) ;
begin
if (champ <> NIL) then
  begin
          champ.tabstop:=En;
          Champ.Enabled :=En;
          Champ.ReadOnly := not En;
          if En then
              champ.color :=ClWindow
          else
              champ.color :=ClBtnFace;
  end;
end;

Function OkLettrage (Where : string; Listb : TListBox; TOBLET : TOB=nil) : Boolean;
var
Q        : TQuery;
Oklettre : Boolean;
Totdebit : double;
Totcredit: double;
Compte   : string;
TL       : TOB;
begin
    Oklettre := TRUE;
// ajout me   Q := OpenSQL ('select e_general,e_auxiliaire,sum(e_debit) TOTDEBIT,sum(e_credit) TOTCREDIT from ecriture  ' + 'Where '+  Where+ ' and e_lettrage<>"" '
    Q := OpenSQL ('select e_general,e_auxiliaire,sum(e_debit) TOTDEBIT,sum(e_credit) TOTCREDIT from ecriture  ' + 'Where e_lettrage<>"" '
       + ' and E_ETATLETTRAGE="TL"'
       +' group by e_general,e_auxiliaire HAVING SUM(E_DEBIT) <> SUM(E_CREDIT)', TRUE);
    while not Q.EOF do
     begin
      Totdebit := Arrondi(Q.FindField ('TOTDEBIT').asFloat,V_PGI.OkDecV);
      Totcredit := Arrondi(Q.FindField ('TOTCREDIT').asFloat ,V_PGI.OkDecV);
      if TOBLET <> nil then
      begin
             if (Q.FindField ('E_AUXILIAIRE').asstring <> '') then
                Compte := Q.FindField ('E_AUXILIAIRE').asstring
             else
                Compte:= Q.FindField ('E_GENERAL').asstring;
             TL := TOBLET.FindFirst(['COMPTE'], [Compte], FALSE);
             if TL <> nil then
             begin
                 Totdebit := Totdebit + Arrondi(TL.getvalue('DEBIT'),V_PGI.OkDecV);
                 Totcredit := Totdebit + Arrondi(TL.getvalue('CREDIT'),V_PGI.OkDecV);
             end;
      end;

      if Totdebit <> Totcredit then
      begin
           Oklettre := FALSE;
           if Q.FindField ('e_auxiliaire').asstring <> '' then
              AfficheListeCom('compte : '+ Q.FindField ('e_auxiliaire').asstring + ' lettrage déséquilibré', Listb)
           else
              AfficheListeCom('compte : '+ Q.FindField ('e_general').asstring+ ' lettrage déséquilibré', Listb)
      end;
      Q.next;
     end;
     ferme(Q);
     Result := Oklettre;
end;

procedure InitDossierEnNombre (Var GrandNbGene, GrandNbAux : Boolean; WhereGene : string=''; WhereAux : string=''); // AJOUT ME 14-01-2005
var
Q1 : TQuery;
begin
    // pb chargement gros dossier + 28000 lignes
    Q1 := OpenSQl ('SELECT COUNT(G_GENERAL) FROM GENERAUX ' + WhereGene,TRUE);
    GrandNbGene := (Q1.Fields[0].asinteger > MaxNombre);
    ferme (Q1);

    Q1 := OpenSQl ('SELECT COUNT(T_AUXILIAIRE) FROM TIERS '+ WhereAux,TRUE);
    GrandNbAux := (Q1.Fields[0].asinteger > MaxNombre);
    ferme (Q1);

end;

Function GetPeriodeJour ( LaDate : TDateTime ) : integer ;
Var YY,MM,DD : Word ;
BEGIN
DecodeDate(LaDate,YY,MM,DD) ;
Result:=100*YY+MM;
Result:= Result*100+DD;
END ;

procedure UpdateComLettrage(What : Byte; Condition : string='');
(*
what 0 : Maj E_IO Pour Flager les ecritures modifiées après lettrage
     1 : Pour reinit de E_PAQUETREVISION
     2 : Pour reinit de E_IO
     3 : Pour reinit de E_PAQUETREVISION Et E_IO
*)
var
  Q                              : TQuery;
  TobMvt                         : TOB;
  St                             : string;
  i                              : Integer;
  WhereSelect, ChampUpdate       : string;
  Date1,Date2                    : TDateTime;
begin
  Date1 := GetParamsoc('SO_CPDATESYNCHRO1');
  Date2 := GetParamsoc('SO_CPDATESYNCHRO2');

  St := '';
  if (Condition <> '') and ((What = 3) or (What = 4) or (What = 5)) then
           St := ' AND '+ Condition;

  case What of
    0:
      begin
        WhereSelect := 'E_PAQUETREVISION=1 ';
        ChampUpdate := 'E_IO="X" ';
        if Condition  <> '' then WhereSelect := WhereSelect + ' AND ' + Condition;

      end;
    1:
      begin
        WhereSelect := 'E_PAQUETREVISION=1 ';
        ChampUpdate := 'E_PAQUETREVISION=0 ';
      end;
    2:
      begin
        WhereSelect := 'E_IO<>"X" ';
        ChampUpdate := 'E_IO="-" ';
      end;
    3:
      begin // E_IO est à X on met à zéro
        WhereSelect := '((E_PAQUETREVISION=1) OR (E_IO="X")) ';
        // fiche 10181
        if ExisteSQl('SELECT E_IO FROM ECRITURE WHERE E_IO="1"' + St) then
        begin
          ExecuteSQL('UPDATE ECRITURE SET E_IO="-" WHERE E_IO="1"' + St);
          SetParamsoc('SO_CPDATESYNCHRO2', iDate1900);
        end;
        // ajout me pour version 590 + case 4 et 5
        if ExisteSQl('SELECT E_IO FROM ECRITURE WHERE E_IO="0"' + St) then
        begin
          ExecuteSQL('UPDATE ECRITURE SET E_IO="1" WHERE E_IO="0"' + St);
          SetParamsoc('SO_CPDATESYNCHRO2', GetParamsoc('SO_CPDATESYNCHRO1'));
        end;
        ChampUpdate := 'E_PAQUETREVISION=0,E_IO="0" ';
        SetParamsoc('SO_CPDATESYNCHRO1', NowH);
      end;
    4:
      begin // E_IO est à 0 on met à 1
        WhereSelect := '(E_PAQUETREVISION=1) OR (E_IO="0") ';
        ChampUpdate := 'E_PAQUETREVISION=0,E_IO="1" ';
        SetParamsoc('SO_CPDATESYNCHRO2', NowH);
        ExecuteSQL('UPDATE ECRITURE SET E_IO="-" WHERE E_IO="1"' + St);
      end;
    5:
      begin // E_IO est à 1 on met à blanc
        WhereSelect := '(E_PAQUETREVISION=1) OR (E_IO="1") ';
        ChampUpdate := 'E_PAQUETREVISION=0,E_IO="-" ';
        SetParamsoc('SO_CPDATESYNCHRO2', iDate1900);
      end;
    6,7 : // annulation synchro
      begin
        WhereSelect := '(E_IO="X"  or E_IO="0"  or E_IO="1")';
        ChampUpdate := 'E_IO="-" ';
        SetParamsoc('SO_CPDATESYNCHRO1', iDate1900);
        SetParamsoc('SO_CPDATESYNCHRO2', iDate1900);
      end;
      8 : // envoi de toute les écritures
      begin
        WhereSelect := '(E_IO<>"X") ';
        ChampUpdate := 'E_IO="X" ';
        SetParamsoc('SO_CPDATESYNCHRO1', iDate1900);
        SetParamsoc('SO_CPDATESYNCHRO2', iDate1900);
      end;

  end;
(* cas de lettrage s'il est fait il faut mettre E_IO à X pour toutes les bordereaux correspondantes *)
if What = 0 then
begin
  TOBMvt := TOB.Create('', nil, -1);
  Q :=
    OpenSQL('SELECT E_EXERCICE,E_JOURNAL,E_PERIODE,E_NUMEROPIECE FROM ECRITURE WHERE '
    + WhereSelect, TRUE);
  TobMvt.LoadDetailDB('ECRITURE', '', '', Q, False, True);
  Ferme(Q);
  for i := 0 to TobMvt.Detail.Count - 1 do
  begin
    St := 'UPDATE ECRITURE SET ' + ChampUpdate + ' WHERE E_EXERCICE="' +
      TobMvt.Detail[i].GetValue('E_EXERCICE') + '" AND E_JOURNAL="' +
      TobMvt.Detail[i].GetValue('E_JOURNAL')
      + '" AND E_PERIODE=' + IntToStr(TobMvt.Detail[i].GetValue('E_PERIODE')) +
        ' '
      + 'AND E_NUMEROPIECE=' +
        IntToStr(TobMvt.Detail[i].GetValue('E_NUMEROPIECE')) + ' ';
    ExecuteSQL(St);
  end;
  TobMvt.Free;
  // fiche 10503
  TOBMvt := TOB.Create('', nil, -1);
  Q := OpenSQL ('SELECT max(E_IO),min(e_io),E_EXERCICE,E_JOURNAL,E_PERIODE,E_NUMEROPIECE,E_QUALIFPIECE FROM ECRITURE GROUP BY E_EXERCICE,E_JOURNAL,E_PERIODE,E_NUMEROPIECE,E_QUALIFPIECE having max(e_io)<>min(e_io) '+
// ajout me pour ne pas envoyer les pièces qui ont dans une même pièce un e_io à 0 et -
  'and (max(e_io)="X" or min(e_io)="X")', TRUE);
  TobMvt.LoadDetailDB('ECRITURE', '', '', Q, False, True);
  Ferme(Q);
  for i := 0 to TobMvt.Detail.Count - 1 do
  begin
    St := 'UPDATE ECRITURE SET ' + ChampUpdate + ' WHERE E_EXERCICE="' +
      TobMvt.Detail[i].GetValue('E_EXERCICE') + '" AND E_JOURNAL="' +
      TobMvt.Detail[i].GetValue('E_JOURNAL')
      + '" AND E_PERIODE=' + IntToStr(TobMvt.Detail[i].GetValue('E_PERIODE')) +
        ' '
      + 'AND E_NUMEROPIECE=' +
        IntToStr(TobMvt.Detail[i].GetValue('E_NUMEROPIECE')) + ' ';
    ExecuteSQL(St);
  end;
  TobMvt.Free;
end
else
begin
  if (What = 6) or (What = 7) then // annulation synchro   on met la reference révision S1 à 0  en dossier complet
  begin
    if (What = 6) then ExecuteSQL('UPDATE ECRITURE SET E_REFREVISION=0 ');
    ExecuteSQL('UPDATE ECRITURE SET E_PAQUETREVISION=0 '); // on remet le paquet revision à 0 pour le lettrage
  end;
  St := 'UPDATE ECRITURE SET ' + ChampUpdate + ' WHERE ' + WhereSelect;
  if Condition <> '' then
  St := St + ' AND '+ Condition;

  if ExisteSQl('SELECT E_IO FROM ECRITURE WHERE ' + WhereSelect) then
  begin
       ExecuteSQL(St);
       if (What = 3) or (What = 4) or (What = 5) then
       begin
            if not ExisteSQl('SELECT E_IO FROM ECRITURE Where E_IO="0"') then
               SetParamsoc('SO_CPDATESYNCHRO1', iDate1900);
            if not ExisteSQl('SELECT E_IO FROM ECRITURE Where E_IO="1"') then
               SetParamsoc('SO_CPDATESYNCHRO2', iDate1900);
    end;
  end
  else
  begin
       if (What = 3) or (What = 4) or (What = 5) then
       begin
            SetParamsoc('SO_CPDATESYNCHRO1', Date1);
            SetParamsoc('SO_CPDATESYNCHRO2', Date2);
       end;
  end;
end;

end;

{$IFDEF EAGLSERVER}

Function RenseigneWrootPath (WhatAppli : string) : string;
begin
       Result := GetFromRegistry(HKEY_LOCAL_MACHINE, GlobalServiceKey, 'RootPath', GlobalModulePath + '\wwwroot', TRUE)+'\'+WhatAppli+'\'+TCPContexte.GetCurrent.MySession.SessionId;
end;

procedure LanceCreatDir (Appli : string);
var
 RepertoireRoot,ReaAppli   : string;
begin
     ReaAppli := GetFromRegistry(HKEY_LOCAL_MACHINE, GlobalServiceKey, 'RootPath', GlobalModulePath + '\wwwroot', TRUE)+'\'+Appli;
     if not DirectoryExists(ReaAppli) then CreateDir(ReaAppli);
     RepertoireRoot := ReaAppli +'\'+TCPContexte.GetCurrent.MySession.SessionId;
     if not DirectoryExists(RepertoireRoot) then CreateDir(RepertoireRoot);
end;

Procedure LanceRemoveDirectory (Appli : string);
Var
Directory : string;
begin
       Directory := GetFromRegistry(HKEY_LOCAL_MACHINE, GlobalServiceKey, 'RootPath', GlobalModulePath + '\wwwroot', TRUE)+'\'+Appli+'\'+TCPContexte.GetCurrent.MySession.SessionId;
       RemoveInDir2 (Directory, TRUE, TRUE, TRUE);
end;

{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : Récupération dans Resultat du numéro contenu dans Buf 
Suite ........ : après épuration des caractères non numériques
Suite ........ : (reprise de la procedure ForceNumerique dans
Suite ........ : ..\paie\lib\PgOutils2.pas)
Mots clefs ... : CHAINE, NUMERIQUE, EPURE
*****************************************************************}
procedure EpureChar (Buf: string; var Resultat: string);
var
  i : integer;

begin
  Resultat  := '';
  for i := 1 to Length(Buf) do
  begin
    if Buf[i] in ['0'..'9'] then
      Resultat  := Resultat + Copy(Buf, i, 1);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 17/04/2007
Modifié le ... :   /  /    
Description .. : reprise de la fonction dans \paie\lib\pgoutils2.pas
Mots clefs ... : 
*****************************************************************}
function SupprimeFichier(FileN: string; Silence: Boolean = False): Boolean;
begin
result := True;
if FileExists(FileN) then
   begin
   if not Silence then
      begin
      if PgiAsk ('Voulez-vous supprimer le fichier '+ExtractFileName(FileN)+'?',
         'Fichier existant :') = MrYes then
         begin
         Result := True;
         DeleteFile(PChar(FileN))
         end
      else
         Result := False;
      end
   else
      begin
      Result := True;
      DeleteFile(PChar(FileN))
      end;
   end;
end;

function PGUpperCase(const S: string; withAccent: boolean = true): string;
var
Ch: Char;
L: Integer;
Source, Dest: PChar;
begin
L:= Length(S);
SetLength(Result, L);
Source:= Pointer(S);
Dest:= Pointer(Result);
while L <> 0 do
      begin
      Ch:= Source^;
      if withAccent then
         Ch:= PGSupprimeaccent (Ch);
      if (Ch >= 'a') and (Ch <= 'z') then
         Dec (Ch, 32);
      Dest^:= Ch;
      Inc (Source);
      Inc (Dest);
      Dec (L);
      end;
end;

function PGSupprimeaccent(Ch: Char): Char;
begin
result:= Ch;
if Ch in ['á', 'à', 'â', 'ã', 'ä', 'å', 'æ'] then
   result:= 'a'
else
if Ch in ['Á', 'À', 'Â', 'Ã', 'Ä', 'Å', 'Æ'] then
   result:= 'A'
else
if Ch in ['ç'] then
   result:= 'c'
else
if Ch in ['Ç'] then
   result:= 'C'
else
if Ch in ['é', 'è', 'ê', 'ë'] then
   result:= 'e'
else
if Ch in ['É', 'È', 'Ê', 'Ë'] then
   result:= 'E'
else
if Ch in ['í', 'ì', 'î', 'ï'] then
   result:= 'i'
else
if Ch in ['Í', 'Ì', 'Î', 'Ï'] then
   result:= 'I'
else
if Ch in ['ñ'] then
   result:= 'n'
else
if Ch in ['Ñ'] then
   result:= 'N'
else
if Ch in ['ó', 'ò', 'ô', 'õ', 'ö'] then
   result:= 'o'
else
if Ch in ['Ó', 'Ò', 'Ô', 'Õ', 'Ö'] then
   result:= 'O'
else
if Ch in ['ú', 'ù', 'û', 'ü'] then
   result:= 'u'
else
if Ch in ['Ú', 'Ù', 'Û', 'Ü'] then
   result:= 'U'
else
if Ch in ['ÿ'] then
   result:= 'y'
else
if Ch in ['Ÿ'] then
   result:= 'Y';
end;

procedure PaysISOLib(Pays: string; var CodeIso, Libelle: string);
var
Q: Tquery;
begin
CodeIso := '';
Libelle := '';
Q:= OpenSql ('SELECT PY_CODEISO2, PY_LIBELLE'+
            ' FROM PAYS WHERE'+
            ' PY_PAYS="'+Pays+'"', True);
if not Q.EOF then
   begin
   CodeIso := Q.FindField('PY_CODEISO2').asString;
   Libelle := Q.FindField('PY_LIBELLE').asString;
   end;
ferme(Q);
end;

end.


