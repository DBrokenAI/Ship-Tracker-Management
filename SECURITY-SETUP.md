# Security setup: Firebase Authentication cutover

Login now uses Firebase Authentication. Passwords are checked by Google, and access to the database is enforced by `firestore.rules`. The old login (client-side hash, `admin123`) is gone.

**Everyone, including the admin, has to "Request Access" once after the cutover.** The old password hashes can't be converted.

## Do these steps in order

1. **Turn on Email/Password sign-in**
   Firebase console → project **ship-checker** → Build → Authentication → Get started → Sign-in method → **Email/Password** → Enable → Save.
   (Leave "Email link" off. Nobody gets emailed. Usernames are mapped to `username@ship-checker.firebaseapp.com` internally.)

2. **Deploy the new `index.html`** (push to GitHub, or Firebase Hosting; see step 7). Open browsers update themselves within about 30 seconds.

3. **Re-create the admin account.** Open the site → Request Access → username `DBrown` + a new strong password.

4. **Promote it to admin by hand** (only needed this once):
   Firebase console → Firestore Database → `users` → open the document whose `username` is `DBrown` **and has no `passwordHash` field** (that one is new) → set `status` = `approved`, `isAdmin` = `true` (boolean).

5. **Publish the security rules**
   Firebase console → Firestore Database → Rules → paste the contents of `firestore.rules` → Publish.
   Or from this folder: `npx firebase-tools login`, then `npx firebase-tools deploy --only firestore:rules`.

6. **Tell the team to Request Access again**, then approve them in 👑 Admin Panel. When everyone is back, click **Delete old records** in the Admin Panel to remove the legacy accounts (they still contain the old password hashes).

7. **(Recommended) Move hosting to Firebase and make the GitHub repo private**
   `npx firebase-tools deploy --only hosting` publishes the site at `https://ship-checker.web.app`.
   Then GitHub → repo Settings → Danger Zone → Change visibility → Private (GitHub Pages stops working on the free plan; that's expected).

## Still temporary: the Outlook VBA macro

The macro doesn't sign in, so the rules still allow, **without login**:
- reading `settings/vbaCode` and `settings/vbaConfig` (so it can self-update)
- creating `incomingEmails` documents (no reading, editing or deleting)

Only admins can **change** the VBA code now, which closes the "anyone can push macros to coworkers" hole.
To finish the lockdown, export the current macro (Admin Panel → Push VBA Update box → copy all) and have it updated to sign in. Then remove the two `TEMPORARY` lines in `firestore.rules`.

**After step 5, check that Outlook still captures emails.** If it stops, the macro is reading or writing something else. Send the `.bas` and the rules can be adjusted.

## Day-to-day
- **Change your password:** sidebar → 🔑 Change Password.
- **Someone forgot their password:** Firebase console → Authentication → Users → delete `theirname@ship-checker.firebaseapp.com`, delete their row in Admin Panel, and they Request Access again.
- **Deactivating or denying** a user in Admin Panel takes effect immediately, both in the app and at the database level.
