## Converting to Powershell. See Powershell folder. This was originally written in Python, but is now converting to Powershell. This will help it run faster and won't require a stand-alone EXE with all of the Python libraries. This is a work in progress.

## BFXDIAGS:  I was a long-time Altiris engineer starting in 2002. After Symantec (Now Broadcom) took over around 2010, someone built a utility a few years later called RAAD (Remote Altiris Agent Diagnostics) for troubleshooting Altiris agents. I am attempting to do the same for BigFix since I work with that product now.

Log viewer built in Python (single EXE) with specific items to troubleshoot the BigFix agent on an endpoint. 

Create and run this file from C:\bfxdiags as the cache files are written there.

* Checks for BESClient Helper status

* Checks for BESClient status

* Checks TCP/UDP ports status on port 52311

* Displays logged on user (app running context)

* Displays Workstation name

* Displays any and all IP addresses assigned to the machine.

* Displays BESClient version

* Displays assigned relays

* Displays Errors in RED, Warnings in  YELLOW, Successes in LIGHT GREEN.

* Restart BESClient service.

* Network speed status (UL and DL)

* Can edit BigFix client registry properties directly.

* Supported Platforms:
These utilities run under Windows.

* License:
Please see the file named LICENSE.

* Issues:
Please submit questions, comments, bugs, enhancement requests.

* Disclaimer:
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
