unit FactGen;
 
interface

Uses HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
{$IFDEF EAGLCLIENT}
     Maineagl,
{$ELSE}
     DB, DBTables, FE_Main,
{$ENDIF}
     SysUtils, Dialogs, Utiltarif, UtilPGI, AGLInit, EntGC, SaisUtil, Windows,
     Forms, Classes, AGLInitGC, UtilArticle , UtilDimArticle, HDimension, HMsgBox,
//uses
      FactUtil ;

// libellés des messages
Function  EcritTOBPiece(NaturePiece,Tiers,Affaire,Etab:String; TOBL:TOB):TOB;
Procedure EcritTOBLigne(NewTOBL, OldTOBL, NewTOBP:TOB; iLigne:Integer);

implementation

Uses
    FactComm, FactGrp,VoirTob;



end.
