unit CalCulContr;

interface
uses
 SysUtils, Classes, UTob, HEnt1, uDbxDataSet, HCtrls, SaisUtil, DB, 
{$IFNDEF EAGLSERVER}
  hmsgbox,
  RapSuppr,
  Dialogs,
{$ENDIF}
 Variants,
 SaisComm,
 HStatus;

Type

      TOBC = Class
           public
            Desc  : HTStrings ;
            T     : T_V ;
            NbC   : integer ;
            Function  GetMvt(Nom : String) : Variant ;
            Procedure ChargeMvt( F : TDataSet ) ;
            Constructor Create(F : TDataSet) ;
            Function TrouveIndice ( T : HTStrings ; Nom : String ; Parle : boolean ) : integer ;
            END ;

  TCalculContr = class
  private
    Exercice, Date1, Date2, Journal1, Journal2, NaturePiece : string;
    LPiece  : TList ;
    procedure LanceMajContreparties;
    procedure TraitePieces;
    procedure ChargePiece (Q : TQuery);
    function  GetGeneTreso(OL : TOBC) : String17 ;
    Procedure MajLigneAna(OL : TOBC ; Gene,Auxi : String17) ;
    Procedure MajLigne(OL : TOBC ; Gene,Auxi : String17) ;
    function  FabriqueRequete : string ;
    Procedure GetCptsContreP (OL1 : TOBC; Lig : Integer ; Var LeGene,LeAux : String ) ;
    procedure FreeListe(T : TList) ;
    procedure ClearListe(T : TList) ;
    Function EstUnTiers (L : TOBC): boolean ;
    Function EstUnCpteHT (L : TOBC) : boolean ;
  public
     constructor Create (TE : Tob);
  end;

procedure LanceTraitementContreparties(RequestTOB : TOB);

var
FCalculContr : TCalculContr;
LDesc : HTStrings ;

implementation

procedure LanceTraitementContreparties(RequestTOB : TOB);
BEGIN
FCalculContr := TCalculContr.Create(RequestTOB);
With FCalculContr do
begin
    LanceMajContreparties;

    Try
     BEGINTRANS ;
     TraitePieces ;
     COMMITTRANS ;
    Except
     on e:exception do
      begin
      {$IFNDEF EAGLSERVER}
       MessageAlerte('Erreur lors du traitement ' + #13#10 + e.message );
      {$ELSE}
          RequestTOB.Putvalue ('ERROR', 'Erreur lors du traitement ' + #13#10 + e.message);
       {$ENDIF}
       RollBack ;
      end;
    End ;
  free;
end;
END ;

procedure TCalculContr.LanceMajContreparties;
BEGIN
Try
 BEGINTRANS ;
 TraitePieces ;
 COMMITTRANS ;
Except
 on e:exception do
  begin
   {$IFNDEF EAGLSERVER}
   MessageAlerte('Erreur lors du traitement ' + #13#10 + e.message );
   {$ENDIF EAGLSERVER}
   RollBack ;
  end;
End ;
END ;

constructor TCalculContr.Create (TE : Tob);
var
L1 : TOB;
begin
     if TE=nil then exit;
     L1                  := TE.Detail.Items[0];
     Exercice            := L1.GetValue('E_EXERCICE') ;
     Date1               := L1.GetValue('E_DATECOMPTABLE') ;
     Date2               := L1.GetValue('E_DATECOMPTABLE_') ;
     Journal1            := L1.GetValue('E_JOURNAL') ;
     Journal2            := L1.GetValue('E_JOURNAL_') ;
     NaturePiece         := L1.GetValue('E_NATUREPIECE') ;
end;


procedure TCalculContr.TraitePieces ;
var OL : TOBC ;
    i : Integer ;
    Q         : TQuery;
    St        : string;
    LeGene, LeAux : string;
BEGIN
St := FabriqueRequete ;
Q := OpenSQl (St, TRUE);
if Q.Eof then BEGIN Ferme(Q) ; Exit ; END ;
LDesc:=HTStringList.Create ;
Q.GetFieldNames(LDesc) ;
LPiece:=TList.Create ;
InitMove(RecordsCount(Q),'') ;
While not Q.Eof do
  BEGIN
  ChargePiece (Q) ;
  //Traitement de la pièce
  i:=0 ;
  While i<=LPiece.Count-1 do
    BEGIN
    OL:=LPiece[i] ;
    GetCptsContreP (OL, i, LeGene, LeAux) ;
    MajLigne(OL, LeGene, LeAux) ;
    Inc(i) ;
    END ;
  END ;
FiniMove ;
Ferme (Q);
LDesc.Clear ; LDesc.Free ;
FreeListe(LPiece) ;
END ;



procedure TCalculContr.ChargePiece (Q : TQuery);
var NumL, NumGrp : Integer ;
    O            : TOBC ;
BEGIN
ClearListe(LPiece) ;
NumL:=0 ; NumGrp := Q.FindField('E_NUMGROUPEECR').AsInteger;
if not Q.EOF and (Q.FindField('E_MODESAISIE').Asstring = 'BOR') then
begin
    While not Q.Eof and (Q.FindField('E_NUMLIGNE').AsInteger>=NumL) and
    (NumGrp = Q.FindField('E_NUMGROUPEECR').AsInteger) do
      BEGIN
      NumL   := Q.FindField('E_NUMLIGNE').AsInteger ;
      NumGrp := Q.FindField('E_NUMGROUPEECR').AsInteger ;
      O:=TOBC.Create(Q) ;
      LPiece.Add(O) ;
      Q.Next ;
      MoveCur(False) ;
      END ;
end
else
begin
    While not Q.Eof and (Q.FindField('E_NUMLIGNE').AsInteger>=NumL) do
      BEGIN
      NumL   := Q.FindField('E_NUMLIGNE').AsInteger ;
      O:=TOBC.Create(Q) ;
      LPiece.Add(O) ;
      Q.Next ;
      MoveCur(False) ;
      END ;
end;
END ;


function TCalculContr.GetGeneTreso(OL : TOBC) : String17 ;
var St : String ;
BEGIN
St:=OL.GetMvt('J_CONTREPARTIE') ;
Result:=TrouveAuto(St,1) ;
if (Result='') then
  if ((OL.GetMvt('G_NATUREGENE')='BQE') or (OL.GetMvt('G_NATUREGENE')='CAI')) then Result:=OL.GetMvt('E_GENERAL') ;
END ;


Procedure TCalculContr.MajLigneAna(OL : TOBC ; Gene,Auxi : String17) ;
var St : String ;
BEGIN
If OL.GetMvt('E_ANA')<>'X' Then Exit ;
St:='UPDATE ANALYTIQ SET Y_CONTREPARTIEGEN="'+Gene+'",'+'Y_CONTREPARTIEAUX="'+Auxi+'" WHERE'
   +' Y_JOURNAL="'+OL.GetMvt('E_JOURNAL')+'" AND Y_DATECOMPTABLE="'
   +USDateTime(OL.GetMvt('E_DATECOMPTABLE'))+'" AND Y_NATUREPIECE="'
   +OL.GetMvt('E_NATUREPIECE')+'" AND Y_NUMEROPIECE='
   +IntToStr(OL.GetMvt('E_NUMEROPIECE'))+' AND Y_NUMLIGNE='+IntToStr(OL.GetMvt('E_NUMLIGNE'))+' AND '
   +' Y_QUALIFPIECE="'+OL.GetMvt('E_QUALIFPIECE')+'" AND Y_EXERCICE="'+OL.GetMvt('E_EXERCICE')+'" ' ;
ExecuteSQL(St) ;
END ;

Procedure TCalculContr.MajLigne(OL : TOBC ; Gene,Auxi : String17) ;
var St : String ;
BEGIN
if (OL.GetMvt('E_CONTREPARTIEGEN')=Gene) and (OL.GetMvt('E_CONTREPARTIEAUX')=Auxi) then Else
  BEGIN
  St:='UPDATE ECRITURE SET E_CONTREPARTIEGEN="'+Gene+'",'+'E_CONTREPARTIEAUX="'+Auxi+'" WHERE'
     +' E_JOURNAL="'+OL.GetMvt('E_JOURNAL')+'" AND E_DATECOMPTABLE="'
     +USDateTime(OL.GetMvt('E_DATECOMPTABLE'))+'" AND E_NATUREPIECE="'
     +OL.GetMvt('E_NATUREPIECE')+'" AND E_NUMEROPIECE='
     +IntToStr(OL.GetMvt('E_NUMEROPIECE'))+' AND E_NUMLIGNE='+IntToStr(OL.GetMvt('E_NUMLIGNE'))+' AND '
     +' E_QUALIFPIECE="'+OL.GetMvt('E_QUALIFPIECE')+'" AND E_EXERCICE="'+OL.GetMvt('E_EXERCICE')+'" ' ;
  ExecuteSQL(St) ;
  END ;
MajLigneAna(OL,Gene,Auxi) ;
END ;

function TCalculContr.FabriqueRequete : string ;
var
St : string;
BEGIN
St := 'SELECT E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NATUREPIECE,E_NUMEROPIECE,E_NUMGROUPEECR,E_MODESAISIE,' ;
St := St + 'E_NUMLIGNE,E_GENERAL,E_AUXILIAIRE,E_CONTREPARTIEGEN,E_CONTREPARTIEAUX, E_TYPEMVT' ;
St := St + ',E_DEBIT,E_CREDIT,G_NATUREGENE,J_NATUREJAL,J_EFFET, J_CONTREPARTIE,E_QUALIFPIECE,E_ANA FROM ECRITURE' ;
St := St + ' LEFT JOIN JOURNAL ON E_JOURNAL=J_JOURNAL' ;
St := St + ' LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL' ;
St := St + ' WHERE J_JOURNAL>="'+Journal1+'"' ;
St := St + ' AND J_JOURNAL<="'+Journal2+'"' ;
St := St + ' AND (J_NATUREJAL="BQE" OR J_NATUREJAL="CAI" OR J_NATUREJAL="ACH" OR J_NATUREJAL="VTE"'+
          ' OR J_NATUREJAL="OD")' ;
if (Exercice<>'') then St := St + ' AND E_EXERCICE="'+Exercice+'"'
                          else St := St + ' AND E_EXERCICE<>""' ;
St := St + ' AND E_DATECOMPTABLE>="'+Date1+'"' ;
St := St + ' AND E_DATECOMPTABLE<="'+Date2+'"' ;
St := St + ' AND (E_QUALIFPIECE="N" OR E_QUALIFPIECE="S" OR E_QUALIFPIECE="R" OR E_QUALIFPIECE="U") ' ;
St := St + ' AND (E_ECRANOUVEAU="N") ' ;
if (Naturepiece<>'') then  St := St + ' AND E_NATUREPIECE="'+Naturepiece+'"' ;
Result := St + ' ORDER BY E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE' ;
END ;

Procedure TCalculContr.GetCptsContreP (OL1 : TOBC; Lig : Integer ; Var LeGene,LeAux : String ) ;
Var k     : integer ;
    Cpte  : string;
    Aux   : string;
    OkBQE : Boolean;
    OkJalEffet : Boolean;
    OL    : TOBC;
    CpteTreso : string;
    procedure Affecte;
    begin
          OL:=LPiece[k] ;
          LeGene:= OL.GetMvt('E_GENERAL') ;
          LeAux := OL.GetMvt('E_AUXILIAIRE') ;
    end;
BEGIN
Cpte    := OL1.GetMvt('E_GENERAL');
Aux     := OL1.GetMvt('E_AUXILIAIRE');
OkBQE   := (OL1.GetMvt('J_NATUREJAL')='BQE') or (OL1.GetMvt('J_NATUREJAL')='CAI') ;
OkJalEffet := OL1.GetMvt('J_EFFET') = 'X' ;
CpteTreso  := GetGeneTreso(OL1);

LeGene:='' ; LeAux:='' ;
if OkBqe then
   BEGIN
   {Pièce de banque}
   if Cpte = CpteTreso then
      BEGIN
      {Sur cpte de banque, contrep=première ligne non banque au dessus}
      for k:=Lig-1 downto 0 do if Cpte <> GetGeneTreso(LPiece[k]) then
          BEGIN
          Affecte;
          Break ;
          END ;
      END else
      BEGIN
      {Sur cpte non banque, contrep=première ligne banque au dessous}
      for k:=Lig+1 to LPiece.Count-1 do if Cpte = GetGeneTreso(LPiece[k]) then
          BEGIN
          Affecte;
          Break ;
          END ;
      END ;
   END else if OkJalEffet then
   BEGIN
   {Pièce sur Journal d'effet}
   if Cpte = GetGeneTreso(LPiece[k]) then
      BEGIN
      {Sur cpte de banque, contrep=première ligne non banque au dessus}
      for k:=Lig-1 downto 0 do if Cpte <>GetGeneTreso(LPiece[k]) then
          BEGIN
          Affecte;
          Break ;
          END ;
      END else
      BEGIN
      {Sur cpte non banque, contrep=première ligne effet au dessous}
      for k:=Lig+1 to  LPiece.Count-1 do if Cpte = GetGeneTreso(LPiece[k]) then
          BEGIN
          Affecte;
          Break ;
          END ;
      END ;
   END else
   BEGIN
      {Cas normal}
      if EstUnTiers(LPiece[Lig]) then
         BEGIN
         {Lecture avant pour trouver "HT"}
         LeAux := '';
         for k:=Lig+1 to  LPiece.Count-1 do if EstUnCpteHT(LPiece[k]) then
         begin
             OL := LPiece[k];
             Affecte;
             Break ;
         end;
         if  LeGene='' then
             BEGIN
             {Lecture arrière pour trouver "HT"}
             for k:=Lig-1 downto 0 do if EstUnCpteHT(LPiece[k]) then
             begin
                 OL := LPiece[k];
                 Affecte;
                 Break ;
             end;
             END ;
         END else
         BEGIN
         {Lecture arrière pour trouver "Tiers"}
        for k:=Lig-1 downto 0 do if EstUnTiers(LPiece[k]) then
         begin
             OL := LPiece[k];
             Affecte;
             Break ;
         end;
         if  LeGene='' then
         BEGIN
             {Lecture avant pour trouver "Tiers"}
             for k:=Lig+1 to  LPiece.Count-1 do if EstUnTiers(LPiece[k]) then
             begin
                 OL := LPiece[k];
                 Affecte;
                 Break ;
             end;
         END ;
      END ;
   END ;
{Cas particuliers}
if LeGene<>'' then Exit ;
if ((OkBqe) and (Cpte <> CpteTreso)) then LeGene:=CpteTreso else
  if not OkBQE then
  BEGIN
     for k:=Lig+1 to LPiece.Count-1 do if GetGeneTreso(LPiece[k]) <> '' then
     begin
      Affecte; Break ;
     end;
     if LeGene='' then for k:=Lig-1 downto 0 do if GetGeneTreso(LPiece[k]) <> '' then
     begin
        Affecte; Break ;
     end;
  END else
  BEGIN
     for k:=Lig-1 downto 0 do
     begin
        OL := LPiece[k];
        if (OL.GetMvt('E_AUXILIAIRE')<>'') then BEGIN Affecte; Break ; END ;
     end;
     if LeGene='' then for k:=Lig+1 to LPiece.Count-1 do
     begin
        OL := LPiece[k];
        if (OL.GetMvt('E_AUXILIAIRE')<>'')then BEGIN Affecte;  Break ; END ;
     end;
  END ;
{Si rien trouvé, swaper les lignes 1 et 2}
if LeGene<>'' then Exit ;
if Lig >=1 then BEGIN OL := LPiece[0]; LeGene:=OL.GetMvt('E_GENERAL') ; LeAux:=Aux ; END else
 if LPiece.Count>=4 then BEGIN OL := LPiece[1]; LeGene:=OL.GetMvt('E_GENERAL') ; LeAux:=Aux ; END ;
END ;

procedure TCalculContr.FreeListe(T : TList) ;
BEGIN
if T=nil then Exit ;
ClearListe(T) ; T.Free ;
END;

procedure TCalculContr.ClearListe(T : TList) ;
var i : Integer ;
BEGIN
for i:=0 to T.Count-1 do TObject(T[i]).Free ;
T.Clear ;
END;

Function TCalculContr.EstUnTiers (L : TOBC): boolean ;
Var Nat : String3 ;
BEGIN
    Result:=True ;
    if (L.GetMvt('E_AUXILIAIRE') <> '') then Exit ;
    if (L.GetMvt('E_GENERAL') <> '') then
    BEGIN
        Nat:=L.GetMvt('G_NATUREGENE') ;
        if ((Nat='TID') or (Nat='TIC')) then Exit ;
    END ;
    Result:=False ;
END ;

Function TCalculContr.EstUnCpteHT (L : TOBC) : boolean ;
Var Nat  : String3 ;
    CC   : string;
BEGIN
    Result := False ;
    if (L.GetMvt('E_AUXILIAIRE') <> '') then Exit ;
    if L.GetMvt('E_TYPEMVT') = 'HT' then Result := True ;
    Nat:=L.GetMvt('G_NATUREGENE') ;
    if ((Nat='CHA') or (Nat='PRO')) then result := True else
    if (Nat='IMO') then  Result := True
    else
    begin
      CC := L.GetMvt('E_GENERAL');
      if (CC[1] = '7') or (CC[1] = '6') then  Result := True ;
    end;
END ;

/// Objet pour mouvements de contreparties

{=====================================================================}
Procedure TOBC.ChargeMvt( F : TDataSet) ;
Var i : Integer ;
begin
if Not F.Active then Exit ; if F.EOF then exit ;
For i:=0 to F.FieldCount-1 do T[i].V:=F.Fields[i].AsVariant ;
end ;

{=====================================================================}
Function TOBC.GetMvt(Nom : String) : Variant ;
Var i : Integer ;
begin
Result:=0 ; i:=TrouveIndice(Desc,Nom,True) ;
if i>=0 then Result:=T[i].V ;
if VarType(Result)=VarNull then Result:='' ;
end ;

{=====================================================================}
Constructor TOBC.Create(F : TDataSet) ;
Var i : Integer ;
begin
Inherited Create ;
if Not F.Active then Exit ; if F.EOF then Exit ;
Desc:=LDesc ; NbC:=LDesc.Count+1   ;
For i:=0 to F.FieldCount-1 do T[i].V:=F.Fields[i].AsVariant ;
end ;

function TOBC.TrouveIndice(T: HTStrings; Nom: String; Parle: boolean): integer;
Var i : integer ;
begin
  Result:=-1 ;
  Nom:=Trim(Nom) ;
  for i:=0 to T.Count-1 do
    if AnsiCompareText(T[i],Nom)=0 then
    	BEGIN
      Result:=i ;
      Break ;
      END ;
  {$IFNDEF EAGLSERVER}
  if ((Result=-1) and (Parle)) then ShowMessage('pas trouvé '+Nom) ;
  {$ENDIF EAGLSERVER}
end;

end.
