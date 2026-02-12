# Lighthouse Demo - Presenter Notes

**Presentation:** Summit Voice AI x Lighthouse (QualifyOS)
**Presenters:** Dan Gill & David Newsom
**Total Time:** ~20-25 minutes + Q&A

---

## Slide 1: Simple Intro
**Time:** ~30 seconds

**What to say:**
> "Thanks for making time. We want to show you what we've been building together. This isn't a pitch deck - this is more of a 'hey, look at this cool thing' kind of conversation."

**Interactive elements:** None - clean intro. Just click to advance.

**Tip:** Don't rush this. Let the slide breathe. The minimal design sets the tone.

---

## Slide 2: The Real Story (Timeline)
**Time:** 2-3 minutes

**What to say at each node:**
- **2020:** "Started Summit as a marketing agency. Learned how businesses actually work from the inside."
- **2022:** "Got obsessed with Voice AI. Built systems the hard way - there were no playbooks, no shortcuts."
- **2023:** "Pitched Aetna on Voice AI for healthcare. They loved it but timing wasn't right. Healthcare stuck with us though."
- **2024:** "A roofing opportunity came along. Built QualifyOS to prove the tech stack works in the real world."
- **2025:** "Met David. Started our Monday/Wednesday/Friday 9:30 PM sessions. That's when everything clicked."
- **Now:** "Building Lighthouse - what insurance support should have been all along."

**Key message:** This wasn't a pivot. It was an evolution. Each step built on the last.

**Transition:** "So let me show you what that roofing proof-of-concept looked like..."

---

## Slide 3: Where We Started - Roofing
**Time:** 1-2 minutes

**What to say:**
> "Before healthcare, we proved everything in roofing. Why? It's the highest-ROI home service vertical and the pain points are crystal clear."

**Walk through bullets:**
- "Every storm creates a surge in demand. The roofer who answers first, wins."
- "We built a 24/7 Voice AI receptionist. Zero missed calls."
- "The whole backend is automated - from lead capture to scheduling to follow-up."

**Point to metrics (right side):**
- "Look at the numbers. 40% missed calls went to zero. 4-hour response time became instant. Conversion nearly tripled."
- "This is what gave us the confidence the stack works."

**Transition:** "So we'd proven the tech. But healthcare was always where we wanted to be..."

---

## Slide 4: What We Saw in Healthcare
**Time:** 2 minutes

**What to say:**
> "Here's what we kept seeing in healthcare."

**Walk through each problem bubble:**
1. "45-minute average wait times. That's just... broken."
2. "$50 million a year in contact center costs for large insurers."
3. "Meanwhile, members are on Reddit and Facebook complaining because nobody's listening."
4. "And the IVR systems? They make it worse. Press 1, press 2, press give up."

**Pause, then read bottom callout:**
> "So we asked ourselves one simple question: What if someone actually listened?"

**Transition:** "That's when we had the idea that changed everything..."

---

## Slide 5: The Idea - Find Them First
**Time:** 2-3 minutes | THIS IS THE BIG IDEA SLIDE

**What to say:**
> "Here's the core insight: don't wait for members to call. Find them where they're already struggling."

**Walk through each node:**
1. **Discovery:** "Every morning at 6 AM, our system scans Reddit, Facebook, LinkedIn, blog comments - anywhere people talk about insurance problems."
2. **Intelligence:** "AI analyzes each post, extracts the pain point, finds contact info, scores urgency on a 1-10 scale."
3. **Response:** "High urgency gets a phone call. Medium gets a text. Lower priority gets an email with resources. Right channel, right time."
4. **Resolution:** "When we connect, it's a real conversation. Not a sales pitch. We help them understand ALL their options and find the best fit."

**Key point to emphasize:**
> "The member never had to call anyone. We found them and helped before they even knew to ask."

**Transition:** "Let me show you what a typical day looks like when this is running..."

---

## Slide 6: How It Actually Works
**Time:** 2-3 minutes

**Walk through each phase (left side):**
- **6 AM:** "System wakes up, starts scanning Reddit, Facebook, LinkedIn, blog comments."
- **7 AM:** "AI processes everything. Extracts the actual struggle, finds contact info, scores urgency."
- **8 AM:** "Outreach begins. High priority gets a phone call. Medium gets a text. Others get a helpful email."
- **All Day:** "Conversations happen naturally. Voice AI talks to members, understands ALL options, finds best fit."

**Point to right panel stats:**
> "These are numbers from our prototype runs. 47 discoveries, 94% resolution rate, under 3 minutes average response."

**Key point:** "This all runs autonomously. No human has to trigger anything."

**Transition:** "Now let me show you what the dashboard actually looks like..."

---

## Slide 7: Live Demo - The Dashboard
**Time:** 3-4 minutes | THIS IS THE MONEY SLIDE

**What to say:**
> "This is what it actually looks like when it's running."

**Walk through sections:**

**KPI Cards (top):**
- "47 discoveries today. 12 active conversations right now. 94% resolution rate."
- "Average response time: 2.3 minutes. Not hours. Minutes."

**Charts (middle):**
- "Pain point distribution shows claim denials are the #1 issue. That tells us where to focus."
- "Sentiment analysis: 62% positive because we're actually helping people."

**Table (bottom):**
- "Here's the magic. Sarah posted on Reddit about a denied claim. We found her at 6 AM, called by 8 AM, resolved by 10."
- "John was confused about coverage on Facebook. Got a personalized text explaining his options."
- Walk through a couple more rows briefly.

**Bottom right corner:**
> "And yeah... this is all running right now."

**Demo tip:** If you have the Replit dashboard open, consider switching to it for a live demo moment here. If not, this slide IS the demo - spend time on it.

**Transition:** "Let me show you the Voice AI that powers those conversations..."

---

## Slide 8: The Voice AI
**Time:** 2-3 minutes

**Walk through capabilities (left side):**
- "24/7. 90+ languages with real-time translation."
- "It's empathetic, not robotic. Trained on actual member conversations."
- "It knows CMS regulations, plan details, all the terminology."
- "Plugs into everything - CRM, claims, billing, EHR."

**Point to conversation preview (right side):**
- Read both messages aloud in a conversational tone.
- "Notice the AI doesn't say 'let me transfer you.' It says 'let me help you right now.'"
- "That's the difference. No hold time. No transfers. Actual help."

**If you have an audio clip of the Voice AI:** Play it here. Even 30 seconds is powerful.

**Key point:** "This isn't a chatbot. It's a knowledgeable, empathetic voice that actually solves problems."

**Transition:** "And this doesn't just work for healthcare..."

---

## Slide 9: Beyond Healthcare
**Time:** 1-2 minutes

**What to say:**
> "Healthcare is our primary focus, but the tech works everywhere."

**Quick touch on each card:**
- "Roofing - that's where we proved it."
- "Home services, insurance brokers, call centers - same pattern applies."
- "Anywhere there's a complex product and frustrated customers, Lighthouse fits."

**Don't linger.** This slide is about showing scale, not deep-diving each vertical.

**Key point:** "The stack is the product. The vertical is just the application."

**Transition:** "Let me show you what's actually under the hood..."

---

## Slide 10: The n8n Build
**Time:** 2-3 minutes

**Walk through the workflow diagram:**
- "Trigger fires at 6 AM every day."
- "Hits Reddit, Facebook, LinkedIn APIs simultaneously."
- "Everything flows into Claude AI for analysis - sentiment, pain points, urgency scoring."
- "Data stores in Google Sheets and CRM."
- "Router splits by priority: high gets a call, medium gets a text, lower gets email."

**Point to callout annotations.**

**Read bottom quote:**
> "With tools like Claude Code, we can prototype workflows like this in hours, not weeks. The speed of iteration is insane."

**Point to tech stack badges:**
> "Our stack: n8n for automation, Claude for intelligence, custom Voice AI, all hosted on Replit."

**Closing:**
> "This is a demo. But it's a working demo. And we're adding to it every single day."
>
> "So... that's what we've been building. What questions do you have?"

---

## General Tips

1. **Pacing:** Let slides breathe. The dark design with glowing elements looks best when you give people a moment to take it in.
2. **Tone:** Casual, confident, slightly nerdy. You built something cool and you're excited about it.
3. **If someone asks a tough question:** "Great question. Let me pull up the dashboard / workflow and show you."
4. **If the tech demo breaks:** "That's live software for you. The point is, this is real and running, not vapor."
5. **Energy:** Start calm (slides 1-3), build energy (slides 4-6), peak at dashboard (slide 7), maintain through close (slides 8-10).
