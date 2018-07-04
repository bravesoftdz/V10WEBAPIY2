unit ZoomEdtProspect;

interface

uses Hent1,sysutils, forms, controls,Hctrls,
     UtofCltVisites,EntGC,uEntCommun;

Procedure RTZoomEdt(Qui : Integer ; Quoi : String) ;
Procedure RTZoomEdtEtat(Quoi : String) ;

implementation
uses {$IFNDEF GCGC} QR, {$ENDIF}
      ZoomEdtGescom ,classes ;

Procedure RTZoomEdtEtat(Quoi : String) ;
var i,Qui: Integer ;
begin
  i := Pos(';',Quoi);
  Qui :=StrToInt (Copy (Quoi,1,i-1));
  Quoi:=Copy(Quoi,i+1,Length(Quoi)-i) ;
  case Qui Of    { valeurs définies dans tablette TTZOOMEDT }
{$IFNDEF GCGC}
    0..89    : ZoomEdt(Qui,Quoi) ;       // COMPTA
{$ENDIF}
    // 90..99   : ZoomEdtImmo(Qui,Quoi) ;   // IMMOS
    300..399 : GCZoomEdt(Qui,Quoi) ;     // GESCOM
    500..599 : RTZoomEdt(Qui,Quoi) ;

  end ;
end;

Procedure RTZoomStringToCleDoc ( StA : String ; Var CleDoc : R_CleDoc ) ;
Var StC : String ;
Begin
StC:=uppercase(Trim(ReadTokenSt(StA))) ; if StC='' then Exit ;
CleDoc.NaturePiece:=StC ;
StC:=uppercase(Trim(ReadTokenSt(StA)))  ; if StC='' then Exit else CleDoc.DatePiece:=UsDateTimeToDateTime(StC) ;
StC:=uppercase(Trim(ReadTokenSt (StA))) ; if StC='' then Exit else CleDoc.Souche:=StC ;
StC:=uppercase(Trim(ReadTokenSt(StA)))  ; if StC='' then Exit else CleDoc.NumeroPiece:=StrtoInt(StC) ;
StC:=uppercase(Trim(ReadTokenSt(StA)))  ; CleDoc.Indice:=StrToInt(StC) ;
end;


Procedure RTZoomEdt(Qui : Integer ; Quoi : String) ;
var st,s : string;
BEGIN
s:=quoi;
Case Qui Of
  500 :  BEGIN      // Tiers
        V_PGI.DispatchTT (28,taConsult ,Quoi, '','');
        END;
// MNG
  501 :  BEGIN      // Contact
        V_PGI.DispatchTT (16,taConsult ,Quoi, '','');
        END;

  502 :  BEGIN      // Action
        V_PGI.DispatchTT (22,taConsult ,Quoi, '','');
        END;
  503 :  BEGIN      // Operation
        st := ReadTokenSt (s);
        s := 'PRODUITPGI='+s;
        V_PGI.DispatchTT (23,taConsult ,st, '',s);
        END;
  504 :  BEGIN      // Commercial
        V_PGI.DispatchTT (9,taConsult ,Quoi, '','');
        END;
  505 :  BEGIN      // Liste des tiers ss actions 
        ZoomListeTiers(Quoi);
        END;
  506 :  BEGIN      // Action  + shift = tiers
        st := ReadTokenPipe(s,';/;') ;
        if (ssShift in V_PGI.FZoomKeyClick) and (s<>'') then
        begin
        V_PGI.DispatchTT (28,taConsult ,s, '','');   //tiers
        end else
        V_PGI.DispatchTT (22,taConsult ,st, '','');   // actions
        END;
  507 :  BEGIN      // Suspect
        V_PGI.DispatchTT (46,taConsult ,Quoi, '','');
        END ;
  END ;
Screen.Cursor:=crDefault ;
END ;

end.
