function StopSignal() {
  BreakPoints.Stop();
}

class BreakPoints {

  static IsStoppedByUser() {
    return BreakPoints.StopByUser;
  }

  static Stop() {
    BreakPoints.StopByUser = true;
  }

  static Reload() {
    BreakPoints.StopByUser = false;
  }
}