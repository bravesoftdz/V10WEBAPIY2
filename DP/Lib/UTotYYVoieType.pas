Unit UTotYYVoieType ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOT ;

Type
  TOT_YYVOIETYPE = Class ( TOT )
    procedure OnUpdateRecord    ; override ;
//    procedure OnNewRecord       ; override ;
  end ;

//////////// IMPLEMENTATION ///////////
Implementation

procedure TOT_YYVOIETYPE.OnUpdateRecord ;
var txt: String;
begin
  Inherited ;
  //--- Remplacer le code suivant par ce qui est en commentaire lors de la maj socref
  txt := GetField('CC_ABREGE');
  SetField('CC_LIBELLE', UpperCase(txt[1])+Copy(txt, 2, Length(txt)-1));
  {
  txt := GetField('YDS_ABREGE');
  SetField('YDS_LIBELLE', UpperCase(txt[1])+Copy(txt, 2, Length(txt)-1));}
end ;

//--- Enlever le commentaire pour mise à jour dans la socref
{
procedure TOT_YYVOIETYPE.OnNewRecord () ;
begin
 Inherited;
 SetField ('YDS_LIBRE','');
 SetField ('YDS_NODOSSIER','000000');
 SetField ('YDS_PREDEFINI','STD');
end ;
}

Initialization
  registerclasses ( [ TOT_YYVOIETYPE ] ) ;
end.

