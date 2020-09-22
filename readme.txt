DS2 Cipher (aka Digitally Secure Encryption)
By: David Greenwood <dsguk@lycos.com>
and David Midkiff <mdj2023@hotmail.com>

Copyright © 2001-2002 David Greenwood and David Midkiff.
All rights reserved.

This algorithm is free for use in any non-commercial project but
you must receive permission from both David Greenwood and David
Midkiff to use this algorithm in commercial projects. Information
on the algorithm can be found in the attached text file or by visiting
our website at http://go.to/ds2cipher.


Version 2.0
-----------

DS2 is incompatible with all previous versions of the DS cipher. We
encourage you to use version 2 as it is faster and much more secure
than before. What's new in DS2?

- The sbox initialization is now effected by the key adding 10x the security.
- The algorithm now loops at 4 rounds providing 4x the security without
  sacrificing much speed in VB. Differential cryptanalysis is much harder
  to achieve with a 4 round structure.
- A secure pseudo-random number generator developed by David Midkiff is now
  fully integrated into the salt generator of DS2 making it much much harder
  to cryptanalyze the algorithm.
- The salt function now adds 6 bytes to the stream providing 247 trillion possible
  ciphertext outputs to a single input.
- Several technical changes have been applied allowing for better error
  correction and a more stable cipher.
- The progress event has been fixed providing faster and smoother updates to 
  the overall progress during encryption and decryption.
