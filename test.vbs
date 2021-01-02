Sub QuickSort(arr, idx)
	dim lstack(512)
	dim rstack(512)
	dim sp
	sp = 1
	dim ltmp, rtmp
	dim l, r, x, swap

	'0番目をpush
	lstack(0) = 0
	rstack(0) = UBound(arr)
	
	while(sp > 0)
		'pop
		sp = sp -1
		ltmp = lstack(sp)
		rtmp = rstack(sp)

		if ltmp < rtmp then
			l = ltmp
			r = rtmp
			x = arr(int((ltmp+rtmp)/2))(idx)

			Do 
				while (arr(l)(idx) > x): l = l+1 :wend
				while (x > arr(r)(idx)): r = r-1 :wend

				if l >= r then
					exit Do
				else
					'swap
					swap = arr(l)
					arr(l) = arr(r)
					arr(r) = swap
					
					if arr(l)(idx)=arr(r)(idx) then
						r = r -1
					end if

				End If
			Loop

			'push
			if (l - ltmp < rtmp -l) then
				lstack(sp) = l + 1
				rstack(sp) = rtmp
				sp = sp + 1
				lstack(sp) = ltmp
				rstack(sp) = l - 1
				sp = sp + 1
			else
				lstack(sp) = ltmp
				rstack(sp) = l - 1
				sp = sp + 1
				lstack(sp) = l + 1
				rstack(sp) = rtmp
				sp = sp + 1
			end if
		end if
	wend
End Sub



Sub QuickSortTest
	dim d
	d = array( _
		array(1, "hogehoge"), _
		array(9, "fugafuga"), _
		array(3, "abaa"), _
		array(4, "dusu") _
	)
	echo d
	
	Call QuickSort(d, 0)
	echo d
End Sub
Sub Echo(arr)
	dim str, c
	str = ""
	for each c in arr
		str = str & c(0) & ":" & c(1) & vbcrlf
	next
	WScript.Echo str
End Sub

QuickSortTest