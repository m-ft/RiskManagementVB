
' Value at Risk 
Function varmean(V, c, sigma, time)
varmean = V * Application.NormSInv(1 - c) * sigma * Sqr(time)
End Function

Function varzero(V, c, sigma, mu, time)
varzero = V * (Application.NormSInv(1 - c) * sigma * Sqr(time) - mu * time)
End Function
