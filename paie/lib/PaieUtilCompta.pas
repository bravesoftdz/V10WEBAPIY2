unit PaieUtilCompta;

interface

uses  Windows,
      StdCtrls,
      Controls,
      Classes,
      sysutils,
      Hctrls,
      UTOB,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,
{$ELSE}
      HQry,
{$ENDIF}
      Hent1;

Type RMVT = RECORD
            Axe,Etabl,Jal,Exo,CodeD,Simul,Nature,LeGuide,TypeGuide : String3 ;
            General,Section : String17 ;
            DateC,DateTaux : TDateTime ;
            Num,LastNumCreat : Longint ;
            TauxD   : Double ;
            Valide,EtapeRegle,FromGuide,ANouveau,SaisieGuidee,ModeOppose : boolean ;
            NumLigne,NumEche,NumLigVisu : integer ;
            Treso,MajDirecte,Effet : Boolean ;
            Souche,FormatCFONB,Document,EnvoiTrans,ModeSaisieJal,GroupeEncadeca,MPGUnique : String3 ;
            Indic  : String[1] ;
            TypeSaisie : String[2] ;
            ExportCFONB,Bordereau,Globalise : boolean ;
            TIDTIC                          : Boolean ;
            ForceModif : Boolean ;
            BloquePieceImport : Boolean ;
            NumEncaDeca : String ;
            End ;

Type TTypeEcr =(EcrGen,EcrBud,EcrGui,EcrAna,EcrAbo,EcrSais,EcrClo) ;

function GetNewNumJal ( Jal : String3; Normale : boolean; DD : TDateTime ; NewMode : Boolean = FALSE ; Cpt : string = '' ; ModeSaisie : string = '' ) : Longint;
FUNCTION QUELEXODT(DD : TDateTime) : String ;
Function SENSENC ( D,C : Double ) : String3 ;

implementation

uses Ent1;
Function SENSENC ( D,C : Double ) : String3 ;
Var Solde : Double ;
BEGIN
SensEnc:='' ;
Solde:=D-C ; if Solde=0 then Exit ;
if Solde>0 then SensEnc:='ENC' else SensEnc:='DEC' ;
END ;


FUNCTION QUELEXODT(DD : TDateTime) : String ;
Var i : Integer ;
begin
Result:=VH^.EnCours.Code ;
If (dd>=VH^.EnCours.Deb) and (dd<=VH^.EnCours.Fin) then Result:=VH^.EnCours.Code else
   If (dd>=VH^.Suivant.Deb) and (dd<=VH^.Suivant.Fin) then Result:=VH^.Suivant.Code Else
      If (dd>=VH^.Precedent.Deb) and (dd<=VH^.Precedent.Fin) then Result:=VH^.Precedent.Code Else
      BEGIN
         For i:=1 To 5 Do
           BEGIN
           If (dd>=VH^.ExoClo[i].Deb) And (dd<=VH^.ExoClo[i].Fin)then BEGIN Result:=VH^.ExoClo[i].Code ; Exit ; END ;
           END ;
      END ;
end ;


(*======================================================================*)
Function OkSoucheN1(DD : TDateTime) : Boolean ;
BEGIN
Result:=(VH^.MultiSouche) And (QuelExoDT(DD)=VH^.Suivant.Code) ;
END ;

(*======================================================================*)
FUNCTION GETNUM (TypeEcr : TTypeEcr ; Facturier : String3 ; Var MasqueNum : String17 ; DD : TDateTime) : LongInt ;
Var Q : TQuery ;
    SQL : String ;
begin
Result:=0 ; MasqueNum:='' ;
Case TypeEcr Of
  EcrGen,EcrAna,EcrClo :
       BEGIN
       SQL:='Select SH_TYPE, SH_SOUCHE, SH_NUMDEPART, SH_SIMULATION, SH_MASQUENUM,SH_NUMDEPARTS, SH_SOUCHEEXO ';
       SQL:=SQL+'FROM SOUCHE WHERE SH_TYPE="CPT" AND SH_SOUCHE="'+Facturier+'"' ;
       Q:=OpenSQL(SQL,TRUE) ;
       if Not Q.Eof then
          BEGIN
          Result:=Q.FindField('SH_NUMDEPART').AsInteger ;
          MasqueNum:=Q.FindField('SH_MASQUENUM').AsString ;
          If OkSoucheN1(DD) And (Q.FindField('SH_SOUCHEEXO').AsString='X') Then Result:=Q.FindField('SH_NUMDEPARTS').AsInteger ;
          END ;
       Ferme(Q) ;
       END ;
  EcrBud :
       BEGIN
       SQL:='Select SH_TYPE, SH_SOUCHE, SH_NUMDEPART, SH_SIMULATION, SH_MASQUENUM ';
       SQL:=SQL+'FROM SOUCHE WHERE SH_TYPE="BUD" AND SH_SOUCHE="'+Facturier+'"' ;
       Q:=OpenSQL(SQL,TRUE) ;
       if Not Q.Eof then
          BEGIN
          Result:=Q.FindField('SH_NUMDEPART').AsInteger ;
          MasqueNum:=Q.FindField('SH_MASQUENUM').AsString ;
          END ;
       Ferme(Q) ;
       END ;
  end ;
end ;

(*======================================================================*)
PROCEDURE SETNUM(TypeEcr : TTypeEcr ; Facturier : String3 ; Num,OldN : Longint ; DD : TDateTime) ;
Var Q : TQuery ;
    OkN1 : Boolean ;
begin
if Num<=0 then BEGIN V_PGI.IoError:=oeUnknown ; Exit ; END ; OkN1:=FALSE ;
Case TypeEcr of
  EcrGen,EcrAna,EcrClo :
    BEGIN
    If OkSoucheN1(DD) Then
      BEGIN
      Q:=OpenSQL('SELECT SH_SOUCHEEXO FROM SOUCHE WHERE SH_TYPE="CPT" AND SH_SOUCHE="'+Facturier+'" ',TRUE) ;
      If Not Q.Eof Then If Q.Fields[0].AsString='X' Then OkN1:=TRUE ; Ferme(Q) ;
      If OkN1 Then
        BEGIN
        If ExecuteSQL('UPDATE SOUCHE SET SH_NUMDEPARTS='+IntToStr(Num)+' WHERE SH_TYPE="CPT" AND SH_SOUCHE="'+Facturier+'" AND SH_NUMDEPARTS='+Inttostr(OldN))<=0 then V_PGI.IoError:=oeUnknown ;
        END Else
        BEGIN
        If ExecuteSQL('UPDATE SOUCHE SET SH_NUMDEPART='+IntToStr(Num)+' WHERE SH_TYPE="CPT" AND SH_SOUCHE="'+Facturier+'" AND SH_NUMDEPART='+Inttostr(OldN))<=0 then V_PGI.IoError:=oeUnknown ;
        END ;
      END Else if ExecuteSQL('UPDATE SOUCHE SET SH_NUMDEPART='+IntToStr(Num)+' WHERE SH_TYPE="CPT" AND SH_SOUCHE="'+Facturier+'" AND SH_NUMDEPART='+Inttostr(OldN))<=0 then V_PGI.IoError:=oeUnknown ;
    END ;
  EcrBud : if ExecuteSQL('UPDATE SOUCHE SET SH_NUMDEPART='+IntToStr(Num)+' WHERE SH_TYPE="BUD" AND SH_SOUCHE="'+Facturier+'" AND SH_NUMDEPART='+Inttostr(OldN))<=0 then V_PGI.IoError:=oeUnknown ;
  end ;
end ;

(*======================================================================*)
PROCEDURE SETINCNUM(TypeEcr : TTypeEcr ; Facturier : String3 ; Var Num : LongInt ; DD : TDateTime) ;
Var MM   : String17 ;
    OldN : Longint ;
begin
Num:=GetNum(TypeEcr,Facturier,MM,DD) ;
OldN:=Num ; Inc(Num) ; SetNum(TypeEcr,Facturier,Num,OldN,DD) ; Dec(Num) ;
end ;

 
function GetNewNumJal ( Jal : String3; Normale : boolean; DD : TDateTime ; NewMode : Boolean = FALSE ; Cpt : string = '' ; ModeSaisie : string = '' ) : Longint;
var
 Q                    : TQuery;
 Souche               : string;
 lYear                : Word ;
 lMonth               : Word ;
 lDay                 : Word ;
 lMaxJour             : integer;
 lStDateDebMois       : string;
 lStDateFinMois       : string;
 lStModeSaisie        : string;
begin

 Result := 0;

 if ( Cpt = '' ) or ( ModeSaisie = '' ) then
  begin // recherche des info du journal

   if Normale then
    Q := OpenSQL ( 'Select J_COMPTEURNORMAL, J_MODESAISIE from JOURNAL Where J_JOURNAL="' + Jal + '"', True )
   else
    Q := OpenSQL ( 'Select J_COMPTEURSIMUL, J_MODESAISIE from JOURNAL Where J_JOURNAL="' + Jal + '"', True ) ;


   if not Q.EOF then
    begin
     Souche        := Q.Fields[0].AsString;
     lStModeSaisie := Q.Fields[1].AsString;
    end
     else
      begin  // le journal n'existe pas -> on sort
       Ferme( Q);
       exit;
      end; // if

   Ferme( Q);

  end
   else
    begin
     // on recupere les param de la fct
     Souche        := Cpt;
     lStModeSaisie := ModeSaisie;
    end; // if


 if Not NewMode then
  begin // garantie le mode de fonctionnement prececedent
   if Souche <> '' then
    SetIncNum ( EcrGen, Souche, Result, DD ) ;
  end
   else
    if lStModeSaisie = '-' then
     begin // mode piece
      if Souche <> '' then
       SetIncNum ( EcrGen, Souche, Result, DD ) ;
     end
      else
       begin // mode bordereau ou libre

        DecodeDate( DD , lYear , lMonth , lDay ) ;

        lMaxJour          := DaysPerMonth( lYear, lMonth) ;
        lStDateDebMois    := USDateTime( EncodeDate( lYear, lMonth, 1) );
        lStDateFinMois    := USDateTime( EncodeDate( lYear, lMonth, lMaxJour));

        Q := OpenSql ( 'select MAX(E_NUMEROPIECE) as N from ECRITURE where E_EXERCICE = "' + QuelExoDt ( DD )  + '" '
                        + 'and E_JOURNAL = "' + Jal + '" and E_DATECOMPTABLE >="' +  lStDateDebMois + '" '
                        + 'and E_DATECOMPTABLE <="' +  lStDateFinMois + '" ', true );

        result            := Q.FindField('N').asInteger + 1;

        Ferme( Q );

       end; // if

end;


end.
 
