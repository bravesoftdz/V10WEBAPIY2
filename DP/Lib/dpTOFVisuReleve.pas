unit dpTOFVisuReleve;

interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     UTob, MaineAGL, UtileAGL,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     PrintDBG, FE_Main,
{$ENDIF}
     forms, sysutils, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOF,
     //VCL
    // Dialogs,
     Graphics,
     // AGL
     Vierge,
     Grids,  // pour TGridDrawState
      // composant AGL
     HTB97,
     HSysMenu,
     HFLabel;

procedure CPLanceFicheVisuRel(Argument : String); // VL 230205 FQ 14757

Type
    {TOF VISURELEVE}
     TOF_VisuReleve = class (TOF)
     private
       BFirst,BPrev,BNext,BLast : TButton ;
       FAlternateColor : TColor ;
       GS : THGrid ;
       LaListe: TStringList ;
       NumFicInt  : integer;
       NomRefOK: Boolean ;
       StCpte: string ;
       procedure BImprimerClick(Sender: TObject);
       procedure BFirstClick(Sender: TObject);
       procedure BPrevClick(Sender: TObject);
       procedure BNextClick(Sender: TObject);
       procedure BLastClick(Sender: TObject);
       procedure GetCellCanvas(ACol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
       procedure RecupListeFichier(ListeChaine: string) ;
       procedure AfficherReleveFichier ;
       procedure AfficherReleveBase(NomRef: string) ;
       Procedure ChangeLeCaption(NewCaption: string) ;
       Procedure PositionneLesFleches ;
       procedure InitGrid ;
     public
       procedure OnLoad ; override ;
       procedure OnArgument (S : String ) ; override ; // VL 230205 FQ 14757
       procedure OnClose ; override ;
     END ;
const   {Grid Visualisation}
        SV_DATE        = 0 ;
        SV_DATEVAL     = SV_DATE + 1 ;
        SV_LIBELLE     = SV_DATEVAL + 1 ;
        SV_DEBIT       = SV_LIBELLE + 1 ;
        SV_CREDIT      = SV_DEBIT + 1 ;
        SV_DEVISE      = SV_CREDIT + 1 ;
        SV_TYPE        = SV_DEVISE + 1 ;


implementation

procedure CPLanceFicheVisuRel(Argument : String);
begin
  AglLanceFiche('CP','RLVVISU','','',Argument);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... :   /  /
Description .. : Convertit une date jjmmaa au format TDATETIME avec gestion an 2000
Mots clefs ... : DATE;CONVERTION;
*****************************************************************}
function ConvertDate(DateChar : string) : TDateTime ;
var Year,Month,Day : word ;
begin
Year  := StrToInt(copy(DateChar,5,2));
if ( Year > 90 )  then Year := Year + 1900 else Year := Year + 2000;
Month:=StrToInt(copy(DateChar,3,2));
Day:=StrToInt(copy(DateChar,1,2));
Result:= EncodeDate(Year,Month,Day);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... :   /  /
Description .. : Convertit les montants ETABAC3 en double
Mots clefs ... : ETEBAC3;MONTANT;CONVERTION;
*****************************************************************}
function ConvertMontant(MontantChar : string;Deci:integer) : Double ;
var MontantFloat : Double;Lettre : char;divi,i : integer ;
begin
Lettre:=MontantChar[14] ;
divi:=1;  MontantFloat :=  1;
for i:=1 to deci do divi := divi * 10;
case Lettre of '{' :       begin Lettre := '0';   end;
               '}' :       begin MontantFloat := -1; Lettre := '0';   end;
               'A'..'I' :  begin Dec(Lettre, 16); end;
               'J'..'R' :  begin MontantFloat := -1; Dec(Lettre, 25); end;
               end;
Result:=MontantFloat*Valeur(copy(MontantChar,1,13)+Lettre)/divi;
end;

{////////////////////////////////////////////}
{                 VisuReleve                 }
{////////////////////////////////////////////}
{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... : 15/05/2000
Description .. : Renvoi le paramètre "TypeArgument" qui se trouve dans la chaine Argument
Mots clefs ... : CHAINE;AGLLANCEFICHE;ARGUMENT
*****************************************************************}
function TrouveArgument(Argument: String;TypeArg : string): string;
var StArgument : string; i,lg : integer;
begin
lg:=Length(TypeArg)-1 ; StArgument := Argument ; i:=Pos(TypeArg,StArgument) ;
if i>0 then
  begin
  system.Delete(StArgument,1,i+lg) ;
  Result:=ReadTokenSt(StArgument);
  end
  else
  Result:='';
end;

// VL 230205 FQ 14757
procedure TOF_VisuReleve.OnArgument (S : String );
begin
  Ecran.Caption := 'Visualisation des relevés';
end;

procedure TOF_VisuReleve.OnLoad;
var FF : TFVierge; BImprimer : TButton; NumFichier,NomRef: string;
begin
inherited ;
LaListe:=TStringList.Create ;
NumFichier:='' ; NomRef:='' ; NumFicInt:=0 ;
if not(Ecran is TFVierge) then Exit ;
FF:=TFVierge(Ecran);
if (FF=Nil) then Exit ;
if V_PGI.NumAltCol=0 then FAlternateColor:=clInfoBk else FAlternateColor:=AltColors[V_PGI.NumAltCol] ;
BImprimer:=TButton(GetControl('BIMPRIMER')) ;  if (BImprimer <> nil ) (*and (not Assigned(BImprimer.OnClick))*) then BImprimer.OnClick:=BImprimerClick;
BFirst   :=TButton(GetControl('BFIRST')) ;    if (BFirst <> nil )    and (not Assigned(BFirst.OnClick))    then BFirst.OnClick:=BFirstClick;
BPrev    :=TButton(GetControl('BPREV')) ;     if (BPrev <> nil )     and (not Assigned(BPrev.OnClick))     then BPrev.OnClick:=BPrevClick;
BNext    :=TButton(GetControl('BNEXT')) ;     if (BNext <> nil )     and (not Assigned(BNext.OnClick))     then BNext.OnClick:=BNextClick;
BLast    :=TButton(GetControl('BLAST')) ;     if (BLast <> nil )     and (not Assigned(BLast.OnClick))     then BLast.OnClick:=BLastClick;
GS:=THGrid(GetControl('GRLV'));
initGrid ;
(* Relevés issus de fichiers *)
RecupListeFichier(TrouveArgument(FF.FArgument,'NOMFICHIER=')) ;
NumFichier:=TrouveArgument(FF.FArgument,'NUMFICHIER=') ;
if NumFichier<>'' then NumFicInt:=StrToInt(NumFichier)-1;
if NumFicInt>=0   then AfficherReleveFichier ;
NomRefOK:=False ;
(* Relevés issus de la Base de données *)
NomRef:=TrouveArgument(FF.FArgument,'NOMREF=') ;
if NomRef<>'' then begin AfficherReleveBase(NomRef) ; NomRefOK:=True ; end ;
PositionneLesFleches ;
end;

procedure TOF_VisuReleve.BImprimerClick(Sender: TObject);
var
  s, R : string;
  Q : TQuery;
{$IFDEF EAGLCLIENT}
  T, F : Tob;
  i : Integer;
{$ENDIF}
  function DecoupeRIB : string;
  begin
    Result := '';
    {Récuperation de la première partie réstante du RIB}
    while (Length(R) > 0) and (R[1] <> ' ') do begin
      Result := Result + R[1];
      System.Delete(R, 1, 1);
    end;

    {Suppression des espaces vides entre les partie du RIB}
    while (Length(R) > 0) and (R[1] = ' ') do
      System.Delete(R, 1, 1);
  end;
begin
  s := 'Compte : ' + StCpte;
  {JP 17/06/04 : certains clients trouvent que le RIB n'est pas très explicite !
                 FQ COMTA 14138 et TRESO 10068}
  R := StCpte;
  Q := OpenSQL('SELECT BQ_LIBELLE FROM BANQUECP WHERE BQ_ETABBQ = "' + DecoupeRIB + '" AND ' +
               'BQ_GUICHET = "' + DecoupeRIB + '" AND BQ_NUMEROCOMPTE = "' + DecoupeRIB + '"', True);
  if not Q.EOF then
    s := Q.FindField('BQ_LIBELLE').AsString + ' (' + StCpte + ')';
  Ferme(Q);

{$IFDEF EAGLCLIENT} // VL 230205 FQ 14757
  SourisSablier;
  T := TOB.Create('non', nil, -1);
  for i := 1 to GS.RowCount-1 do begin
    F := TOB.Create('non', T, -1);
    F.AddChampSup('DATE', False);
    F.AddChampSup('DATEV', False);
    F.AddChampSup('LIBELLE', False);
    F.AddChampSup('DEBIT', False);
    F.AddChampSup('CREDIT', False);
    F.AddChampSup('DEVISE', False);
    F.AddChampSup('TITRE', False);
    F.SetString('TITRE', S);
    F.SetString('DATE', GS.Cells[0, i]);
    F.SetString('DATEV', GS.Cells[1, i]);
    F.SetString('LIBELLE', GS.Cells[2, i]);
    F.SetString('DEBIT', GS.Cells[3, i]);
    F.SetString('CREDIT', GS.Cells[4, i]);
    F.SetString('DEVISE', GS.Cells[5, i]);
  end;
  SourisNormale;
  LanceEtatTob('E','CST','RLV',T, True, False, False, nil, '', Ecran.Caption, False, 0, '', 0, '');
  T.Free;
{$ELSE}
  PrintDBGrid(GS, Nil, s, '') ;
{$endif}
End;

procedure TOF_VisuReleve.GetCellCanvas(ACol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
begin
  if (GS.Cells[SV_TYPE,ARow]='1') or (GS.Cells[SV_TYPE,ARow]='7') then begin
    GS.Canvas.Font.color:=clBlack;
    GS.Canvas.Brush.Color:=FAlternateColor;
    end ;
end;

procedure TOF_VisuReleve.OnClose ;
begin
LaListe.free ;
end ;

procedure TOF_VisuReleve.RecupListeFichier(ListeChaine: string) ;
begin
repeat LaListe.Add(ReadTokenStV(ListeChaine)) ; until ListeChaine='' ;
end ;

procedure TOF_VisuReleve.AfficherReleveFichier ;
var Montant : double ;Deci : integer;  FSource : Textfile ; StReleve,CodeOpe,NomFichier : string;
begin
NomFichier:='' ; InitGrid ;
if (NumFicInt>=0) and (NumFicInt<=LaListe.Count-1 ) then NomFichier:=LaListe.Strings[NumFicInt] ;
if NomFichier=''  then Exit ;
AssignFile(FSource,NomFichier);
Reset(FSource); { Taille d'enregistrement = 1 }
try
  While Not eof(FSource) do Begin
    Readln(FSource,StReleve) ;
    CodeOpe := copy(StReleve,1,2) ;
    if (CodeOpe = '01') or (CodeOpe = '04') or (CodeOpe = '07') then begin
      Deci := StrToInt(copy(StReleve,20,1));
      Montant:=ConvertMontant(copy(StReleve,91,14),Deci);
      GS.Cells[SV_DATE,GS.RowCount-1]:=DateToStr(ConvertDate(copy(StReleve,35,6))) ;
      if (Montant<0) then GS.Cells[SV_DEBIT,GS.RowCount-1] := FloatToStrF(Montant*-1,ffFixed,20,Deci);
      if (Montant>0) then GS.Cells[SV_CREDIT,GS.RowCount-1] := FloatToStrF(Montant,ffFixed,20,Deci);
      if (Montant=0) then GS.Cells[SV_CREDIT,GS.RowCount-1] := '0.00' ;
      GS.Cells[SV_DEVISE,GS.RowCount-1] := copy(StReleve,17,3);
      if (CodeOpe = '01')  then begin
        StCpte:=copy(StReleve,3,5)+' '+copy(StReleve,12,5)+' '+copy(StReleve,22,11) ;
        ChangeLeCaption('Compte : '+StCpte) ;
        GS.Cells[SV_LIBELLE,GS.RowCount-1]:='Solde Initial ';{Libelle}
        GS.Cells[SV_TYPE,GS.RowCount-1] := '1';{}
        end;
      if (CodeOpe = '04') then begin
        GS.Cells[SV_LIBELLE,GS.RowCount-1] := copy(StReleve,49,31);{Libelle}
        GS.Cells[SV_DATEVAL,GS.RowCount-1] := DateToStr(ConvertDate(copy(StReleve,43,6))) ;{Date}
        GS.Cells[SV_TYPE,GS.RowCount-1] := '4';{}
        end;
      if (CodeOpe = '07') then
        begin
        GS.Cells[SV_LIBELLE,GS.RowCount-1] := 'Solde Final';{Libelle}
        GS.Cells[SV_TYPE,GS.RowCount-1] := '7';{}
        end ;
      GS.RowCount:=GS.RowCount+1;
      end;
    end ;
finally
  CloseFile(FSource);
end ;
GS.RowCount:=GS.RowCount-1;
end ;

procedure TOF_VisuReleve.AfficherReleveBase(NomRef: string);
var Q: TQuery ;
begin
InitGrid ;
if NomRef='' then Exit ;
Q:=OpenSQL('SELECT * FROM EEXBQ WHERE EE_REFPOINTAGE="'+NomRef+'"',TRUE) ;
try
  with GS,Q do begin
  if Eof then begin PgiBox('Impossible d''afficher le relevé','Affichage des relevés') ; Exit ; end ;
  Cells[SV_DATE,   RowCount-1]:=FindField('EE_DATESOLDE').AsString ;
  Cells[SV_DEBIT,  RowCount-1]:=FloatToStrF(FindField('EE_NEWSOLDEDEB').AsFloat,ffFixed,20,2) ;
  Cells[SV_CREDIT, RowCount-1]:=FloatToStrF(FindField('EE_NEWSOLDECRE').AsFloat,ffFixed,20,2) ;
  Cells[SV_DEVISE, RowCount-1]:=FindField('EE_DEVISE').AsString ;
  Cells[SV_LIBELLE,RowCount-1]:='Solde Initial ';
  Cells[SV_TYPE,   RowCount-1]:='1';
  ChangeLeCaption('Compte : '+FindField('EE_RIB').AsString) ;
  RowCount:=RowCount+1 ;
  // Solde Final
  Cells[SV_DATE,   RowCount-1]:=FindField('EE_DATEOLDSOLDE').AsString ;
  Cells[SV_DEBIT,  RowCount-1]:=FloatToStrF(FindField('EE_OLDSOLDEDEB').AsFloat,ffFixed,20,2) ;
  Cells[SV_CREDIT, RowCount-1]:=FloatToStrF(FindField('EE_OLDSOLDECRE').AsFloat,ffFixed,20,2) ;
  Cells[SV_DEVISE, RowCount-1]:=FindField('EE_DEVISE').AsString ; ;
  Cells[SV_LIBELLE,RowCount-1]:='Solde Final ';
  Cells[SV_TYPE,   RowCount-1]:='7';
  end ;
finally
  ferme(Q) ;
  end ;

Q:=OpenSQL('SELECT * FROM EEXBQLIG WHERE CEL_REFPOINTAGE="'+NomRef+'"',TRUE) ;
try
  while not Q.Eof do with GS,Q do begin
    InsertRow(rowcount-1) ;
    Cells[SV_DATE,   RowCount-2]:=FindField('CEL_DATEOPERATION').AsString ;
    Cells[SV_DATEVAL,RowCount-2]:=FindField('CEL_DATEVALEUR').AsString ;
    if FindField('CEL_DEBIT').AsFloat<>0 then Cells[SV_DEBIT, RowCount-2]:=FloatToStrF(FindField('CEL_DEBIT').AsFloat,ffFixed,20,2)
                                         else Cells[SV_CREDIT,RowCount-2]:=FloatToStrF(FindField('CEL_CREDIT').AsFloat,ffFixed,20,2);
    Cells[SV_DEVISE, RowCount-2]:=Cells[SV_DEVISE,1] ;
    Cells[SV_LIBELLE,RowCount-2]:=FindField('CEL_LIBELLE').AsString ;
    Cells[SV_TYPE,RowCount-2]:='4';
    Next ;
    RowCount:=RowCount+1 ;
    end ;
finally
  ferme(Q) ;
  end ;
end ;

procedure TOF_VisuReleve.BFirstClick(Sender: TObject);
begin
NumFicInt:=0 ;
AfficherReleveFichier ;
PositionneLesFleches ;
end ;

procedure TOF_VisuReleve.BPrevClick(Sender: TObject);
begin
dec(NumFicInt) ;
AfficherReleveFichier ;
PositionneLesFleches ;
end ;

procedure TOF_VisuReleve.BNextClick(Sender: TObject);
begin
inc(NumFicInt) ;
AfficherReleveFichier ;
PositionneLesFleches ;
end ;

procedure TOF_VisuReleve.BLastClick(Sender: TObject);
begin
NumFicInt:=LaListe.Count-1 ;
AfficherReleveFichier ;
PositionneLesFleches ;
end ;

Procedure TOF_VisuReleve.ChangeLeCaption(NewCaption: string) ;
var FF : TFVierge;
begin
FF := TFVierge(Ecran);
if (FF<>Nil) then FF.Caption:=NewCaption ;
end ;

Procedure TOF_VisuReleve.PositionneLesFleches ;
begin
if NomRefOk then begin
  if BFirst<>nil then BFirst.Enabled:=False ;
  if BLast<>nil then BLast.Enabled:=False ;
  end
  else begin
  if NumFicInt=0 then begin
    if BFirst<>nil then BFirst.Enabled:=False ;
    if BLast<>nil then BLast.Enabled:=True ;
    end
    else begin
    if NumFicInt=LaListe.Count-1 then begin
      if BFirst<>nil then BFirst.Enabled:=True ;
      if BLast<>nil then BLast.Enabled:=False ;
      end
      else begin
      if BFirst<>nil then BFirst.Enabled:=True ;
      if BLast<>nil then BLast.Enabled:=True ;
      end ;
    end;
  end ;
if (BFirst<>nil) and (BPrev<>nil) then BPrev.Enabled:=BFirst.Enabled ;
if (BLast<>nil) and (BNext<>nil) then BNext.Enabled:=BLast.Enabled ;
end ;

procedure TOF_VisuReleve.InitGrid ;
var i,j : integer ;
begin
if (GS=nil) or not (GS is THGrid) then exit ;
GS.RowCount := 2; GS.ColWidths[SV_TYPE]:=0 ;
GS.GetCellCanvas:=GetCellCanvas ;
GS.FColAligns[SV_DEBIT]:=taRightJustify;
GS.FColAligns[SV_CREDIT]:=taRightJustify;
for j:=1 to GS.RowCount-1 do for i:=0 to GS.ColCount-1 do GS.Cells[i,j]:= '';
end;

Initialization
  registerclasses ( [TOF_VisuReleve]);


end.
