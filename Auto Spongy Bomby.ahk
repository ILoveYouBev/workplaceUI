CoordMode,Mouse,Screen
bunsX := 956
BunsY := 611
PattyX := 1198
PattyY := 648
Grill1X := 1109
Grill1Y :=478
Grill2X := 1120
Grill2Y :=529
Grill3X :=1210
Grill3Y :=482
Grill4X :=1223
Grill4Y :=531
LettuceX :=1016
LettuceY :=455
TomatoX := 1027
TomatoY:=516
OnionX :=609
OnionY :=441
SausageX :=603
SausageY :=507
Pizza1X :=726 
Pizza1Y :=457
Pizza2X := 715
Pizza2Y :=531
ChickenX := 342
ChickenY := 714
ChickenPlaceX := 359
ChickenPlaceY := 604
FriesX := 432
FriesY := 710
FriesPlaceX := 448
FriesPlaceY := 603
ColaX := 250
ColaY :=415 
IceCreamX := 211 
IceCreamY := 656

return

3::
loop,4
	Click(PattyX,PattyY)
loop, 4 
	Click(BunsX,BunsY)
loop,4
{
	Click(LettuceX,LettuceY)
	Click(TomatoX,TomatoY)
}
Click(Grill1X,Grill1Y)
Click(Grill2X,Grill2Y)
Click(Grill3X,Grill3Y)
Click(Grill4X,Grill4Y)

return


1::
loop,2{
	Click(OnionX,OnionY)
	Click(SausageX,SausageY)
}
Click(Pizza1X,Pizza1Y)
Click(Pizza2X,Pizza2Y)

return


2::
Click(FriesPlaceX,FriesPlaceY)
Click(ChickenPlaceX,ChickenPlaceY)
Click(FriesX,FriesY)
Click(ChickenX,ChickenY)

return



Click(WhereX,WhereY){
	Random,RandomPause,25,50
	sleep,RandomPause
	Random,RandomX,-5,5
	Random,RandomY,-5,5
	X := WhereX + RandomX
	Y := WhereY + RandomY
	MouseClick,Left,X,Y
	return
}

return


^Esc::Reload
