unit UtilOperation;

interface
uses Hctrls,Hent1
{$IFDEF EAGLCLIENT}
      ,MaineAGL
{$ELSE}
      ,Fe_Main
{$ENDIF}
      ;

Procedure RTAttachePieceOper(stOperation : String) ;
Procedure RTSuppressionLienPieceOper(stOperation,StNaturePiece,StSouche,
             StNumero,stIndice : String);

implementation

procedure RTAttachePieceOper(stOperation : String) ;
var RefPiece,stNature,stSouche,stNumero,stIndice : string;
begin
  RefPiece:=AGLLanceFiche('RT', 'RTATTACHPIECEOPER', '', '', '');
  if RefPiece <> '' then
     begin
     stNature := ReadTokenSt (RefPiece);
     stSouche := ReadTokenSt (RefPiece);
     stNumero := ReadTokenSt (RefPiece);
     stIndice := ReadTokenSt (RefPiece);
     ExecuteSql('UPDATE PIECE SET GP_OPERATION="'+stOperation+'", '+
       'GP_DATEMODIF="'+UsTime(NowH)+'", '+
       'GP_UTILISATEUR="'+V_PGI.User+'" Where GP_NATUREPIECEG="'+stNature+'"'+
       ' and GP_SOUCHE ="'+ stSouche+'"'+
       ' and GP_NUMERO='+ stNumero+
       ' and GP_INDICEG='+ stIndice);
     end;
end;

Procedure RTSuppressionLienPieceOper(stOperation,StNaturePiece,StSouche,
             StNumero,stIndice : String);
begin
   ExecuteSql('UPDATE PIECE SET GP_OPERATION="",'+ 
     'GP_DATEMODIF="'+UsTime(NowH)+'", '+
     'GP_UTILISATEUR="'+V_PGI.User+'" Where GP_NATUREPIECEG="'+stNaturePiece+'"'+
     ' and GP_SOUCHE ="'+ stSouche+'"'+
     ' and GP_NUMERO='+ stNumero+
     ' and GP_INDICEG='+ stIndice);
end;

end.
