# cg-staff-face-api — selfie check for CG Time Attendance (experimental)

An AWS Lambda function that CG Time Attendance calls to set up and check an
employee's selfie at clock in / clock out. It uses **Amazon Rekognition**
(face collection `cg-staff`, Sydney region `ap-southeast-2`).

This file is kept here so the code is version-controlled; GitHub Pages never
runs it. It is copied into AWS by the setup steps below.

## How it decides who is who

- The app sends the employee's Microsoft **sign-in (ID) token** with every
  request. The function checks the token's signature and that it was issued
  by Concept Group's tenant **to the CG Time Attendance app**.
- The employee's identity comes from that token, never from the phone. Each
  face record is labelled with their sign-in name (`@` written as `:`).
  Nobody can set up or check a face for someone else.
- **First use:** the employee's first selfie becomes their face record.
  It's refused if the face already belongs to another employee.
- **Every clock in/out:** the selfie must match the signed-in employee, and
  must be their best match. Otherwise the app blocks the clock in/out.
- **Reset:** only the people listed in `ADMIN_USERS` can list or reset
  face records (from **Face admin** at the bottom of the app).
- Photos are never stored, only Rekognition face prints.

| Action | Who | Does |
|---|---|---|
| `status` | anyone signed in | Is my face set up? Am I an admin? |
| `enrol` | anyone signed in | Set up my face (once only) |
| `verify` | anyone signed in | Does this selfie match me? |
| `list` | admins | Everyone with a face record |
| `reset` | admins | Remove someone's face record so they can set it up again |

**Limits to know (experimental):** like the GPS gate, this is a strong
deterrent and an audit trail, not a guarantee. It does not yet include
liveness detection (someone holding up a photo of a colleague). Rekognition
offers Face Liveness for that at extra cost if it's needed later.

## Settings (Lambda environment variables)

| Name | Value |
|---|---|
| `TENANT_ID` | Concept Group's tenant ID (as in `shared/cg-config.js`) |
| `CLIENT_ID` | CG Time Attendance's client ID (as in `shared/cg-config.js`) |
| `COLLECTION_ID` | `cg-staff` |
| `ADMIN_USERS` | Comma-separated admin sign-in names, e.g. `mscully@concept.com.pg` |
| `MATCH_THRESHOLD` | Optional, default `90` (minimum similarity %) |

## One-time setup in AWS CloudShell

Open **CloudShell** in the AWS console (Sydney region) and run each command
on its own. Each one prints a short result.

**1. Face collection**
```
aws rekognition create-collection --collection-id cg-staff --region ap-southeast-2
```

**2. A role the function runs as.** It can only write logs and use the `cg-staff` collection.
```
ACCOUNT=$(aws sts get-caller-identity --query Account --output text); echo $ACCOUNT
```
```
aws iam create-role --role-name cg-staff-face-api-role --assume-role-policy-document '{"Version":"2012-10-17","Statement":[{"Effect":"Allow","Principal":{"Service":"lambda.amazonaws.com"},"Action":"sts:AssumeRole"}]}'
```
```
aws iam attach-role-policy --role-name cg-staff-face-api-role --policy-arn arn:aws:iam::aws:policy/service-role/AWSLambdaBasicExecutionRole
```
```
aws iam put-role-policy --role-name cg-staff-face-api-role --policy-name cg-staff-rekognition-access --policy-document "{\"Version\":\"2012-10-17\",\"Statement\":[{\"Effect\":\"Allow\",\"Action\":[\"rekognition:SearchFacesByImage\",\"rekognition:IndexFaces\",\"rekognition:DeleteFaces\",\"rekognition:ListFaces\"],\"Resource\":\"arn:aws:rekognition:ap-southeast-2:$ACCOUNT:collection/cg-staff\"},{\"Effect\":\"Allow\",\"Action\":\"rekognition:DetectFaces\",\"Resource\":\"*\"}]}"
```

**3. The function**, using the code from this folder on `main`.
```
curl -sSo index.mjs https://raw.githubusercontent.com/ConceptGroupPNG/CG_HUBS/main/aws/cg-staff-face-api/index.mjs && zip -q function.zip index.mjs && ls -l function.zip
```
```
aws lambda create-function --function-name cg-staff-face-api --runtime nodejs22.x --handler index.handler --zip-file fileb://function.zip --role arn:aws:iam::$ACCOUNT:role/cg-staff-face-api-role --timeout 15 --memory-size 256 --region ap-southeast-2 --environment "Variables={TENANT_ID=ea9e105c-2df6-47ef-a8e3-b4aa1d788423,CLIENT_ID=04c72756-916b-4fed-aa36-a6771ea4e8dc,COLLECTION_ID=cg-staff,ADMIN_USERS=mscully@concept.com.pg}"
```
If this says the role "cannot be assumed", wait 10 seconds and run it again;
new roles take a moment to become usable.

**4. Its web address.** Only the CG HUBS site may call it from a browser.
```
aws lambda create-function-url-config --function-name cg-staff-face-api --auth-type NONE --region ap-southeast-2 --cors '{"AllowOrigins":["https://conceptgrouppng.github.io"],"AllowMethods":["POST"],"AllowHeaders":["authorization","content-type"],"MaxAge":86400}'
```
```
aws lambda add-permission --function-name cg-staff-face-api --statement-id FunctionURLAllowPublicAccess --action lambda:InvokeFunctionUrl --principal "*" --function-url-auth-type NONE --region ap-southeast-2
```
```
aws lambda add-permission --function-name cg-staff-face-api --statement-id FunctionURLInvokeAllowPublicAccess --action lambda:InvokeFunction --principal "*" --invoked-via-function-url --region ap-southeast-2
```

The sign-in token check (in the code) is what actually protects it. Auth type
NONE only means AWS lets the request reach that check.

**5. Test it.** Without a sign-in, it must refuse:
```
URL=$(aws lambda get-function-url-config --function-name cg-staff-face-api --region ap-southeast-2 --query FunctionUrl --output text); echo $URL
curl -s -X POST $URL -H "content-type: application/json" -d '{"action":"status"}'
```
Expected: `{"error":"Missing sign-in token","code":"UNAUTHORISED"}`.

Then put the URL into `faceApi` for the `attendance` app in `shared/cg-config.js`
(by pull request). With `faceApi: ""` the selfie check is switched off.

## Updating the code later

After a change to `index.mjs` is merged:
```
curl -sSo index.mjs https://raw.githubusercontent.com/ConceptGroupPNG/CG_HUBS/main/aws/cg-staff-face-api/index.mjs && zip -q function.zip index.mjs && aws lambda update-function-code --function-name cg-staff-face-api --zip-file fileb://function.zip --region ap-southeast-2 --query LastUpdateStatus
```
