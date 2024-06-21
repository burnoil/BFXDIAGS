## Transistor Replacement Data System (TRDS): Python app developed to help with vintage semiconductor replacement. This is primarily used for vintage audio restorations.

## BFXDIAGS:  I was a long-time Altiris engineer since 2002. After Symantec (Now Broadcom) took over around 2010, someone built a utility a few years later called RAAD (Remote Altiris Agent Diagnostics) for troubleshooting Altirtis agents. I am attempting to do the same for BigFix since I work with that product now.

## Log viewer built in Python (single EXE) with specific items to troubleshoot the BigFix agent on an endpoint. This has some specific checks for properties in our environment so those won't work in yours. Most other features will still work, however.

*  Checks for BESClient Helper status

* Checks for BESClient status

* Checks TCP/UDP ports status on port 52311

* Displays logged on user (app running context)

* Displays Workstation name

* Displays any and all IP addresses assigned to the machine.

* Displays BESClient version

* Displays assigned relays

* Has remote agent management ability (still being refined)

* Displays the count of Errors or Warnings in the logs

* Displays Errors in RED, Warnings in  YELLOW, Successes in LIGHT GREEN. (Can be changed by user)

* Button to find and open the latest BigFix log in the default log location (Auto-tails)

* Button to restart BESClient service.

* Button to clear site cache

* Button to clear Agent cache

* Manual port checking to check any port (UDP or TCP)

* Button to open Windows Event Viewer

* Network speed status (UL and DL)

* Button to open and edit BigFix registry tree directly (Double-click to display/edit)

* Most features are threaded.

Adding  more...

* Supported Platforms:
These utilities run under Windows.

* License:
Please see the file named LICENSE.

* Issues:
Please submit questions, comments, bugs, enhancement requests.

* Disclaimer:
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
