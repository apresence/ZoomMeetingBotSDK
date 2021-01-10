namespace ZoomMeetingBotSDK.ControlBot
{
    public class ControlBotInitParam
    {
        public IHostApp hostApp;
    }

    public interface IControlBot
    {
        /// <summary>
        /// Called to initialize the control bot.
        /// </summary>
        void Init(ControlBotInitParam param);

        /// <summary>
        /// Called to start the bot.
        /// </summary>
        void Start();

        /// <summary>
        /// Called once the bot will no longer be used.
        /// </summary>
        void Stop();
    }
}