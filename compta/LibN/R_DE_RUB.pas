unit R_DE_RUB;

interface

uses XVISingleton, HEnt1 ;

Type TRDR = class(TXVISingleStringList)
      Private
       Procedure ChargeRDR ;
       Function GereLesRDR (Formule : hstring) : Variant ;
      protected
       Class Function IsSingleton : Boolean ; Override ;
      public
        PrmAvecAno,PrmEtab,PrmDevi,PrmSDate,PrmCollectif : String ;
        Function GereRDR ( CodeRDR : String) : Variant ;
        function IndexOf(const S: string): Integer; override;
        procedure DoFreeInstance ; override ;
      End ;

Function RubriquedeRubrique : TRDR ;

implementation

uses Ed_Tools,
     HCtrls,
     Formule,
     Ent1,
{$IFNDEF CMPGIS35}
     uLibCalcEdtCompta,
{$ENDIF}
{$IFDEF eAGLClient}
     UTob
{$ELSE}
    {$IFNDEF DBXPRESS} dbtables {$ELSE} uDbxDataSet {$ENDIF}
{$ENDIF}
     ;

type TRDRDetail = Class
      public
       Code,
       Signe,
       Formule : String
      End ;

//////////////////////////////////////////////////////////////////////////////////
Function RubriquedeRubrique : TRDR ;
Begin
   Result:=TRDR.Create ;
End ;
//////////////////////////////////////////////////////////////////////////////////
function TRDR.IndexOf(const S: string): Integer;
Begin
   if Count<=0 then
      ChargeRDR ;
   Result:=Inherited IndexOf(S) ;
End ;
//////////////////////////////////////////////////////////////////////////////////
procedure TRDR.DoFreeInstance ;
BEgin
   VideStringliste(Self) ;
   Inherited ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Class Function TRDR.IsSingleton : Boolean ;
Begin
   Result:=TRUE ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function TRDR.GereLesRDR (Formule : hstring) : Variant ;
Begin
{$IFNDEF CMPGIS35}
   Result:=CalcOleEtat('GETCUMUL','RUB;'+Formule+';'+PrmAvecAno+';'+PrmEtab+';'+PrmDevi+';'+PrmSDate+';'+PrmCollectif+';')
{$ENDIF}
End ;
//////////////////////////////////////////////////////////////////////////////////
Function TRDR.GereRDR( CodeRDR : String) : Variant ;
var
   Idx           : Integer ;
   DD            : TRDRDetail ;
   oldbByGenEtat : Boolean ;
Begin
   Result:=0 ;
   Idx:=indexOf(CodeRDR) ;
   if Idx>-1 then begin
      {$IFDEF NOVH}
      {$ELSE}
      oldbByGenEtat:=VH^.bByGenEtat ;
      {$ENDIF NOVH}
      DD:=TRDRDetail(Objects[Idx]) ;
      Result:=GFormule('{'+DD.Formule+'}',GereLesRDR,nil,1) ;
      if DD.Signe='NEG' then
         Result:=-Result ;
      {$IFDEF NOVH}
      {$ELSE}
      VH^.bByGenEtat:=oldbByGenEtat ;
      {$ENDIF NOVH}
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Procedure TRDR.ChargeRDR ;
var
  QQ : TQuery ;
  DD : TRDRDetail ;
Begin
   QQ:=OpenSQL('Select * from RUBRIQUE where RB_CLASSERUB="RDR"',TRUE) ;
   try
     while Not QQ.eof do Begin
        DD:=TRDRDetail.create ;
        With DD do begin
           Code:=QQ.FindField('RB_RUBRIQUE').AsString ;
           Signe:=QQ.FindField('RB_SIGNERUB').AsString ;
           Formule:=QQ.FindField('RB_COMPTE1').AsString ;
        End ;
        Addobject(DD.Code,DD) ;
        QQ.Next ;
     end ;
   finally
    Ferme (QQ);
   end;
End ;
//////////////////////////////////////////////////////////////////////////////////
end.
