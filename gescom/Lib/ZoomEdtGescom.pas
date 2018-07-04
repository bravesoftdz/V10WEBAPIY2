unit ZoomEdtGescom;

interface

uses Hent1,sysutils, forms, controls, Factutil, FactComm, Facture,
{$IFNDEF GCGC}
     QR,
{$ENDIF}
{$IFDEF GRC}
     ZoomEdtProspect,
{$ENDIF}
    MenuSpec,        //Pour CGEP
    Hctrls, UtilPGI, EntGC,uEntCommun ;

Procedure GCZoomEdt(Qui : Integer ; Quoi : String) ;
Procedure GCZoomEdtEtat(Quoi : String) ;

implementation

Procedure GCZoomEdtEtat(Quoi : String) ;
var i,Qui: Integer ;
begin
  i := Pos(';',Quoi);
  Qui :=StrToInt (Copy (Quoi,1,i-1));
  Quoi:=Copy(Quoi,i+1,Length(Quoi)-i) ;
  if (Qui>=900) and (Qui<=999) and vmDisp_ZoomEdt( Qui,Quoi) then Exit ;   //Pour CGEP
  case Qui Of    { valeurs définies dans tablette TTZOOMEDT }
{$IFNDEF GCGC}
    0..89    : ZoomEdt(Qui,Quoi) ;       // COMPTA
{$ENDIF}
    // 90..99   : ZoomEdtImmo(Qui,Quoi) ;   // IMMOS
    300..399 : GCZoomEdt(Qui,Quoi) ;     // GESCOM
{$IFDEF GRC}
    500..599 : RTZoomEdt(Qui,Quoi) ;
{$ENDIF}
  end ;
end;

Procedure GCZoomStringToCleDoc ( StA : String ; Var CleDoc : R_CleDoc ) ;
Var StC : String ;
Begin
StC:=uppercase(Trim(ReadTokenSt(StA))) ; if StC='' then Exit ;
CleDoc.NaturePiece:=StC ;
StC:=uppercase(Trim(ReadTokenSt(StA)))  ; if StC='' then Exit else CleDoc.DatePiece:=UsDateTimeToDateTime(StC) ;
StC:=uppercase(Trim(ReadTokenSt (StA))) ; if StC='' then Exit else CleDoc.Souche:=StC ;
StC:=uppercase(Trim(ReadTokenSt(StA)))  ; if StC='' then Exit else CleDoc.NumeroPiece:=StrtoInt(StC) ;
StC:=uppercase(Trim(ReadTokenSt(StA)))  ; CleDoc.Indice:=StrToInt(StC) ;
end;


Procedure GCZoomEdt(Qui : Integer ; Quoi : String) ;
Var CleDoc : R_CleDoc;
BEGIN

Case Qui Of
  300:  BEGIN      // article
        V_PGI.DispatchTT (7,taConsult ,Quoi, '','');
        END;
  301:  BEGIN      // tiers
        V_PGI.DispatchTT (8,taConsult ,Quoi, '','');
        END;
  302:  BEGIN      // Pièces
        FillChar(CleDoc,Sizeof(CleDoc),#0) ;
        GCZoomStringToCleDoc(Quoi,CleDoc) ;
        SaisiePiece(CleDoc,taConsult) ;
        END;
  303:  BEGIN      // commercial
        V_PGI.DispatchTT (9,taConsult ,Quoi, '','');
        END;
  304:  BEGIN      // catalogue
        V_PGI.DispatchTT (11,taConsult ,Quoi, '','');
        END;
  305:  BEGIN      // Clients Mode
        V_PGI.DispatchTT (8,taConsult ,Quoi, '','');
        END;
  306:  BEGIN      // Fournisseurs Mode
        V_PGI.DispatchTT (12,taConsult ,Quoi, '','');
        END;
  307:  BEGIN      // Contacts
        V_PGI.DispatchTT (16,taConsult ,Quoi, '','');
        END;
  308:  BEGIN      // Etablissements
        V_PGI.DispatchTT (18,taConsult ,Quoi, '','');
        END;

  END ;
Screen.Cursor:=crDefault ;
END ;

end.
