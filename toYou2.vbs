dim str1
set objTTS  = createobject("sapi.spvoice")
str1="1喜你成疾，药石无医。 2情不知所起，一往而深。 3任凭弱水三千，只取一瓢饮。 4衣带渐宽终不悔，为伊消得人憔悴。 5生死契阔，与子成说。执子之手，与子偕老。 6两情若是久长时，又岂在朝朝暮暮。 7相思相见知何日，此时此夜难为情。 8有美人兮，见之不忘。一日不见兮，思之如狂。 9山有木兮木有枝，心悦君兮君不知。"
objTTS.speak str1


