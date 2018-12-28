# 파이썬과 xingAPI



### 실행환경

- Windows 10

- Miniconda -- version : 4.5.11

- Miniconda 설치 : Python 3.7 & Windows 64-bit (exe installer)

  ​


### API 사용시 주의사항

- API를 사용하기 위해서는 32-bit python 버전을 사용해야 합니다.

  설치한 아나콘다(미니콘다)가 64-bit python을 사용할 경우 새로운 가상환경을 생성하여 진행하셔야 합니다.

  ```
  set CONDA_FORCE_32BIT = 1
  conda create -n [가상환경 이름] python=3.5
  activate [가상환경 이름]
  pip install pywin32 (win32com 모듈 없을 시 설치)
  ```


- API를 사용하여 로그인하기 위해서는 공인인증서가 연결되어 있어야 합니다.

  (USB에 공인인증서가 저장되어 있는 경우, USB 연결 후 진행)

  ​

### 계좌 정보 파일 양식(passwords.csv)

```
type,account_num,user_id,pwd,cert_pwd,url
transaction,계좌번호,로그인 아이디,로그인 비밀번호,공인인증 비밀번호,url
```
