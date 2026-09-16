/** โครงข้อมูลที่ถอดมาจาก JSON ของ TSP_Monitor_API.gs */
#pragma once
#include <Arduino.h>

struct Checkpoint {
  char  t[6];        // "HH:MM"
  long  amount;      // บาท
  int   machines;
};

struct SchedRow {
  char name[18];
  char cell[7][6];   // ค่ากะแต่ละวัน จ-อา เช่น "D","N","OFF"
};

struct MonitorData {
  bool  valid = false;
  char  hash[12]  = "";
  char  ts[18]    = "--:--";
  char  branch[10]= "";
  char  branchName[52] = "";
  int   machines  = 0;

  // วันนี้
  long  dayRev = 0, dayTgt = 0;
  float dayPct = -1;            // -1 = ไม่มีเป้า
  int   dayLogs = 0;

  // เช็คพอยต์
  Checkpoint cp[6];
  int   cpCount = 0;
  long  cash = 0, cashDiff = 0;
  bool  cashFlag = false;

  // สัปดาห์
  long  week[7] = {0,0,0,0,0,0,0};
  long  weekSum = 0, weekPrev = 0;
  float weekChg = 0;
  bool  hasWeekChg = false;

  // เดือน
  char  period[10] = "";
  long  mRev = 0, mTgt = 0, mFore = 0;
  float mPct = -1;
  char  mStat = 'N';            // G / Y / R / N
  int   mMilestone = 0, mElapsed = 0, mDays = 30;
  bool  mShared = false;

  // ตารางกะ
  SchedRow sched[6];
  int   schedRows = 0;

  char  warn[120] = "";         // รวม warning เป็นข้อความเดียว
};
